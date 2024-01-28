import os
import fnmatch
import subprocess
import pandas
import gspread
from subprocess import check_call

# colors
x  = '\033[0m'  # reset
gr = '\033[90m' # grey
lr = '\033[91m' # light red
lg = '\033[92m' # light green
ly = '\033[93m' # light yellow
mg = '\033[95m' # magenta
cy = '\033[96m' # cyan
w  = '\033[97m' # white

test_so_directory = ''
so_pattern = '*.so'
bin_directory = 'C:\\msys64\\usr\\local\\vitasdk\\arm-vita-eabi\\bin\\'
read_elf_exe = 'readelf.exe'
read_elf_path = bin_directory + read_elf_exe
objdump_exe = 'objdump.exe'
objdump_path = bin_directory + objdump_exe
findstr_exe = 'findstr.exe'
findstr_jc_string = 'Java_'
findstr_opensles_strings = ['SL_IID_ANDROIDEFFECT','SL_IID_ANDROIDEFFECTCAPABILITIES', 'SL_IID_ANDROIDEFFECTSEND', 'SL_IID_ANDROIDCONFIGURATION', 'SL_IID_ANDROIDSIMPLEBUFFERQUEUE']

files = os.listdir(test_so_directory)

so_files = fnmatch.filter(files, so_pattern)

data_frame_list = []

for so_file in so_files:
	print(f"Checking {lg}{so_file}{x}")

	game_information = {}
	game_information["game_name"] = ""
	game_information["so_file"] = so_file

	read_elf_arg = test_so_directory + so_file

	output = subprocess.Popen([read_elf_path, "-d", read_elf_arg], stdout=subprocess.PIPE).communicate()[0]

	output_list = output.split(b"\r\n")
	
	needed_lib_list = []

	for unfiltered_lib in output_list:
		if b"NEEDED" in unfiltered_lib:
			unfiltered_lib_string = str(unfiltered_lib)
			lib_string = unfiltered_lib_string[unfiltered_lib_string.find('[')+len('['):unfiltered_lib_string.rfind(']')]

			needed_lib_list.append(lib_string)
	
	if needed_lib_list:
		print(f"...Found the following {mg}NEEDED{x} libs")
		game_information["so_file_needed_libs"] = "\r\n".join(needed_lib_list)
		for lib_string in needed_lib_list:
			print(f"......{lib_string}")
	else:
		game_information["so_file_needed_libs"] = " "

	# Check for JavaCom
	check_call_string = objdump_path + " -T -C " + read_elf_arg + " | findstr " + findstr_jc_string
	script = f"{check_call_string}"
	val = subprocess.Popen(['powershell', '-Command', script], stdout=subprocess.PIPE).communicate()[0]

	filtered_javacom_list = []
	val_list = val.split(b"\r\n")
	for unfiltered_javacom in val_list:
		unfiltered_javacom_string = str(unfiltered_javacom)
		filtered_javacom = unfiltered_javacom_string[unfiltered_javacom_string.find(findstr_jc_string):-1]
		if filtered_javacom:
			filtered_javacom_list.append(filtered_javacom)

	game_information["so_file_found_java_count"] = len(filtered_javacom_list)
	if len(filtered_javacom_list):
		print(f"...Found {mg}{len(filtered_javacom_list)} Java_com{x} functions")

	game_information["so_file_found_java"] = "\r\n".join(filtered_javacom_list)
	for javacom in filtered_javacom_list:
		print(f"......{cy}{javacom}{x}")

	# check for OpenSLES
	if "opensles" in read_elf_arg.lower():
		found_opensles_symbols = []
		for findstr_opensles_string in findstr_opensles_strings:
			check_call_string = objdump_path + " -T -C " + read_elf_arg + " | findstr " + findstr_opensles_string
			script = f"{check_call_string}"
			opensles_val = subprocess.Popen(['powershell', '-Command', script], stdout=subprocess.PIPE).communicate()[0]

			opensles_val_string = str(opensles_val)
			filtered_opensles_val_string = opensles_val_string[opensles_val_string.find(findstr_opensles_string):-1]

			if(filtered_opensles_val_string):
				found_opensles_symbols.append(findstr_opensles_string)
				print(f"...Found {lr}{findstr_opensles_string}{x} symbol")
		game_information["open_sles_found"] = "\r\n".join(found_opensles_symbols)
	else:
		game_information["open_sles_found"] = " "

	data_frame_list.append(game_information)

# create the panda frame with the information and save it to a spreadsheet
panda_dataFrame = pandas.DataFrame(data_frame_list)
panda_dataFrame.loc[len(panda_dataFrame)] = " "
panda_dataFrame.to_excel("output.xlsx", sheet_name="apk_information")

# save the information in the google sheet
gc = gspread.service_account()
sh = gc.open("APK Information")
worksheet = sh.get_worksheet(0)

worksheet.update([panda_dataFrame.columns.values.tolist()] + panda_dataFrame.values.tolist())

# clean up the spreadsheet
rows = worksheet.get_all_values()

# iterate through the rows to see if we need to clean up.
start_row = 0
start_row_value = None
start_match = 0
end_row = 0
current_row_value = None
current_row_count = 0

# iterate through, ignore first row
for row in rows:
	if start_row_value == None:
		start_row_value = row[0]
		continue

	current_row_value = row[0]
	current_row_count = current_row_count + 1

	print(f"Start Row {start_row_value} ... Current Row {current_row_value}")

	if start_row_value == current_row_value and start_match == 0 and len(start_row_value) > 0:
		start_row = current_row_count
		start_match = 1

	if start_row_value != current_row_value and start_match == 1:
		end_row = current_row_count
		start_match = 0

		print(f"Merging rows ... {start_row} through {end_row}")

		worksheet.merge_cells(f"A{start_row}:A{end_row}","MERGE_COLUMNS")

	start_row_value = current_row_value