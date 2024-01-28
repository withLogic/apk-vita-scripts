import os
import fnmatch
import math
import shutil
import subprocess
import pandas
import gspread
from zipfile import ZipFile

# setup some variables
source_folder = 'D:\\ApkTest\\'
apk_pattern   = '*.apk'
so_pattern    = '*.so'

bin_directory = 'C:\\msys64\\usr\\local\\vitasdk\\arm-vita-eabi\\bin\\'
read_elf_exe  = 'readelf.exe'
read_elf_path = bin_directory + read_elf_exe
objdump_exe   = 'objdump.exe'
objdump_path  = bin_directory + objdump_exe
findstr_exe   = 'findstr.exe'
findstr_jc_string = "Java_"
findstr_opensles_strings = ['SL_IID_ANDROIDEFFECT','SL_IID_ANDROIDEFFECTCAPABILITIES', 'SL_IID_ANDROIDEFFECTSEND', 'SL_IID_ANDROIDCONFIGURATION', 'SL_IID_ANDROIDSIMPLEBUFFERQUEUE']

total_apk = 0
total_unity = 0
data_frame_list = []

# borrowed from somewhere on StackOverflow
def convert_size(size_bytes):
	if size_bytes == 0:
		return "0B"
	size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
	i = int(math.floor(math.log(size_bytes, 1024)))
	p = math.pow(1024, i)
	s = round(size_bytes / p, 2)

	return "%s %s" % (s, size_name[i])

# colors
x  = '\033[0m'  # reset
gr = '\033[90m' # grey
lr = '\033[91m' # light red
lg = '\033[92m' # light green
ly = '\033[93m' # light yellow
mg = '\033[95m' # magenta
cy = '\033[96m' # cyan
w  = '\033[97m' # white

files = os.listdir(source_folder)

apk_files = fnmatch.filter(files, apk_pattern)

for apk_file in apk_files:
	apk_path = os.path.join(source_folder, apk_file)
	apk_file_name = os.path.basename(apk_path).split('.')[0]

	print(f"Processing APK: {lg}{apk_path}{x}")

	possible_port_has_armv7 = 0
	possible_port_has_unity = 0
	possible_port_has_gdx = 0
	possible_port_has_glesv3 = 0
	total_apk = total_apk + 1

	try:
		game_information = {}
		game_information["game_name"] = apk_file

		with ZipFile(apk_path) as zip_archive:
			za_lib_list = []
			extracted_so_file = None

			for file_name in zip_archive.namelist():
				if file_name[:3] == 'lib':
					za_file_information = zip_archive.getinfo(file_name)
					za_file_size = convert_size(za_file_information.file_size)

					za_lib_list.append(file_name)

					# pull out any libunity files
					if 'libunity' in file_name:
						print(f"...{lr}{file_name}...{za_file_size}{x}")
						possible_port_has_unity = 1

					# pull out any libgdx files
					elif 'libgdx' in file_name:
						print(f"...{lr}{file_name}...{za_file_size}{x}")
						possible_port_has_gdx = 1

					else:
						print(f"...{cy}{file_name}...{za_file_size} {x}")

					# for any armv7 games, let's extract the .so files
					if 'armeabi-v7a' in file_name:
						possible_port_has_armv7 = 1

						if not os.path.exists(source_folder + apk_file_name):
							os.makedirs(source_folder + apk_file_name)

						extracted_so_file = file_name.split('/')[-1]
						extracted_so_path = source_folder + apk_file_name
						extracted_so_full_path = source_folder + apk_file_name + "\\" + extracted_so_file

						with zip_archive.open(file_name) as zf, open(extracted_so_full_path, 'wb') as f:
							shutil.copyfileobj(zf, f)

			game_information["libs"] = "\r\n".join(za_lib_list)

			#  le's do some deeper examination of the extracted .so files
			if extracted_so_file:
				files = os.listdir(extracted_so_path)
				so_files = fnmatch.filter(files, so_pattern)

				for so_file in so_files:
					print(f"Checking {lg}{so_file}{x}")
					so_file_information = {}

					so_file_information["so_file"] = so_file

					read_elf_arg = extracted_so_path + "\\" + so_file

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
						so_file_information["so_file_needed_libs"] = "\r\n".join(needed_lib_list)
						for lib_string in needed_lib_list:
							print(f"......{lib_string}")
					else:
						so_file_information["so_file_needed_libs"] = " "

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

					so_file_information["so_file_found_java_count"] = len(filtered_javacom_list)
					if len(filtered_javacom_list):
						print(f"...Found {mg}{len(filtered_javacom_list)} Java_com{x} functions")

					so_file_information["so_file_found_java"] = "\r\n".join(filtered_javacom_list)
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
						so_file_information["open_sles_found"] = "\r\n".join(found_opensles_symbols)
					else:
						so_file_information["open_sles_found"] = " "

					so_file_information.update(game_information)
					data_frame_list.append(so_file_information)
			else:
				game_information["so_file"] = " "
				game_information["so_file_needed_libs"] = " "
				game_information["so_file_found_java_count"] = " "
				game_information["so_file_found_java"] = " "
				game_information["open_sles_found"] = " "
				data_frame_list.append(game_information)

	except Exception as error:
		print(f"{error}")
		pass

	if possible_port_has_unity:
		total_unity = total_unity + 1

	if possible_port_has_armv7 and not possible_port_has_unity and not possible_port_has_glesv3 and not possible_port_has_gdx:
		print(f"...{lg}POSSIBLE PORT{x}")
	else:
		print(f"...{lr}UNABLE TO BE PORTED{x}")

# create the panda frame with the information and save it to a spreadsheet
panda_dataFrame = pandas.DataFrame(data_frame_list)
panda_dataFrame.loc[len(panda_dataFrame)] = " "
panda_dataFrame.to_excel("output.xlsx", sheet_name="apk_information")

percent = round((total_unity / total_apk) * 100, 2)

print(f"Total: {total_unity} (unity) / {total_apk} (total) ... {percent}%")

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