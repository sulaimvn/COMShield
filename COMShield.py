import os
import subprocess
import ctypes
import re
import csv
import shutil
import winreg
import time
import sys
from colorama import init
import datetime
import argparse



class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GREY = '\033[90m'
    ENDC = '\033[0m'

WARNING = Colors.YELLOW + "WARNING: " + Colors.ENDC
ERROR = Colors.RED + "ERROR: " + Colors.ENDC



csv_rows = []
rawcopy_path = os.path.join(os.path.dirname(__file__), "RawCopy-master", "RawCopy.exe")
source_usrclass_pattern = r"C:\Users\{}\AppData\Local\Microsoft\Windows\UsrClass.dat"
source_ntuser_pattern = r"C:\Users\{}\NTUSER.DAT"
destination_directory = os.path.join(os.path.dirname(__file__), "output", "userprofile")
clear_dir = os.path.join(os.path.dirname(__file__), "output")
recmd_path = os.path.join(os.path.dirname(__file__), "RECmd", "RECmd.exe")
output_csv_file = os.path.join(os.path.dirname(__file__))



devnull = open(os.devnull, 'w')
def copy_user_profile_data():
    rawcopy_path = os.path.join(os.path.dirname(__file__), "RawCopy-master", "RawCopy.exe")
    source_usrclass_pattern = r"C:\Users\{}\AppData\Local\Microsoft\Windows\UsrClass.dat"    
    source_ntuser_pattern = r"C:\Users\{}\NTUSER.DAT"
    destination_directory = os.path.join(os.path.dirname(__file__), "output", "userprofile")
    devnull = open(os.devnull, 'w')

    for user in os.listdir(r"C:\Users"):

        source_usrclass_file = source_usrclass_pattern.format(user)
        source_ntuser_file = source_ntuser_pattern.format(user)
        usrclass_destination_file = os.path.join(destination_directory, user, "UsrClass.dat")
        ntuser_destination_file = os.path.join(destination_directory, user, "NTUSER.DAT")
        os.makedirs(os.path.dirname(usrclass_destination_file), exist_ok=True)
        os.makedirs(os.path.dirname(ntuser_destination_file), exist_ok=True)
        subprocess.run([rawcopy_path, "/FileNamePath:" + source_usrclass_file, "/OutputPath:" + os.path.dirname(usrclass_destination_file), "/OutputFileName:" + "UsrClass.dat"], stdout=devnull, stderr=subprocess.STDOUT)
        subprocess.run([rawcopy_path, "/FileNamePath:" + source_ntuser_file, "/OutputPath:" + os.path.dirname(ntuser_destination_file), "/OutputFileName:" + "NTUSER.DAT"], stdout=devnull, stderr=subprocess.STDOUT)

        try:
            for user_dir in os.listdir(destination_directory):
                user_dir_path = os.path.join(destination_directory, user_dir)
                usrclass_file_path = os.path.join(user_dir_path, "UsrClass.dat")
                ntuser_file_path = os.path.join(user_dir_path, "NTUSER.DAT")
                if not os.path.exists(usrclass_file_path) or not os.path.exists(ntuser_file_path):
                    for file in os.listdir(user_dir_path):
                        file_path = os.path.join(user_dir_path, file)
                        os.remove(file_path)
                    os.rmdir(user_dir_path)
                    
        except OSError as e:
            if e.errno == 39:
                ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            else:
                raise e
devnull.close()



def get_sids():
    reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList")
    num_subkeys = winreg.QueryInfoKey(reg_key)[0]
    sids = []

    for i in range(num_subkeys):
        sid = winreg.EnumKey(reg_key, i)
        if re.match("^S-1-5-21-\\d+-\\d+-\\d+-\\d+$", sid):
            sids.append(sid)
    winreg.CloseKey(reg_key)

    return sids



def run_recmd(sid):
    output_directory = os.path.join(os.path.dirname(__file__), "output", "userprofile")

    for user in os.listdir(output_directory):
        user_directory = os.path.join(output_directory, user)
        input_file_path = os.path.join(user_directory, "UsrClass.dat")
        output_file_path = os.path.join(user_directory, sid, "UsrClass_CLSID.txt")
        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
        command = '{} -f "{}" --kn "{}_Classes\CLSID" --nl'.format(recmd_path, input_file_path, sid)
        subprocess.run(command, shell=True, stdout=open(output_file_path, "w"), text=True)
        
        with open(output_file_path, "r") as f:
            output = f.read()

        clsid_pattern = re.compile(r"\{?[0-9A-F]{8}-?[0-9A-F]{4}-?[0-9A-F]{4}-?[0-9A-F]{4}-?[0-9A-F]{12}\}?", re.IGNORECASE)
        clsid_values = clsid_pattern.findall(output)
        clsid(clsid_values, input_file_path, destination_directory, sid, user)
        
            
def clsid(clsid_values, input_file_path, destination_directory, sid, user):
    data_regex = re.compile(r"Data:\s+([^\s\(\)]+)")
    data_dict = {}

    for clsid in clsid_values:
        iProcsvr32_command = r"{} -f {} --kn {}_Classes\CLSID\{}\InProcServer32 --nl".format(recmd_path, input_file_path, sid, clsid)
        output_clsid_dir = os.path.join(destination_directory, user, sid, clsid)
        os.makedirs(output_clsid_dir, exist_ok=True)
        output_clsid = os.path.join(output_clsid_dir, "CLSID.txt")
        output = subprocess.run(iProcsvr32_command, shell=True, stdout=subprocess.PIPE).stdout.decode()
        match = data_regex.search(output)
        
        if match:
            data = match.group(1)
            data_dict[clsid] = data
        
    output_csv_file = os.path.join(os.path.dirname(__file__), "allusers_output.csv")
    
    with open(output_csv_file, "a", newline="") as f:
        writer = csv.writer(f) 
        for clsid, data in data_dict.items():
            writer.writerow([user, clsid, data])
            

           
def get_hklm():
    key_path_local = r"Software\Classes\CLSID"
    key_local = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path_local, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
    subkey_names_local = []
    i = 0
    
    while True:
        try:
            subkey_name_local = winreg.EnumKey(key_local, i)
            subkey_names_local.append(subkey_name_local)
        except WindowsError as e:
            if e.winerror == 259:  
                break
        i += 1
    winreg.CloseKey(key_local)
    values_dict_local = {}
    
    for subkey_name in subkey_names_local:
        key_path_local = fr"Software\Classes\CLSID\{subkey_name}\InProcServer32"
        try:
            key_local = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path_local, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        except FileNotFoundError as e:
            continue
        except PermissionError as e:
            continue

        try:
            num_values = winreg.QueryInfoKey(key_local)[1]
            if num_values > 0:
                i = 0
                while True:
                    try:
                        value_name, value_data, value_type = winreg.EnumValue(key_local, i)
                        if value_name == "":
                            value_name = "Default"
                        values_dict_local[f"{subkey_name}_{value_name}"] = value_data
                    except WindowsError as e:
                        if e.winerror == 259:
                            break
                    i += 1
        except Exception as e:
            print(ERROR + time.ctime() + f"\tError accessing key {key_path_local}: {e}")
        finally:
            winreg.CloseKey(key_local)
            
    output_csv_file = os.path.join(os.path.dirname(__file__), "hklm_output.csv")   
    pattern = r"(\{[0-9A-F-]+\})_Default"
    with open(output_csv_file, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(['CLSID', 'Data'])

        for key, value in values_dict_local.items():
            matches = re.search(pattern, key)
            if matches:
                clsid = matches.group(1)
                clsid_key = clsid + "_Default"
                extracted_value = values_dict_local.get(clsid_key)
                writer.writerow([clsid, extracted_value])



def compare_registry_values(fname):
    all_users_file = os.path.join(os.path.dirname(__file__), "allusers_output.csv")
    hklm_file = os.path.join(os.path.dirname(__file__), "hklm_output.csv")
    final_csv = os.path.join(os.path.dirname(__file__), fname)
    all_file_contents = {}
    hklm_contents = {}
   
    with open(hklm_file, "r", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 2:
                clsid = row[0]
                data = row[1]
                hklm_contents[clsid] = data

    differences = []
    with open(all_users_file, "r", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 3:
                user = row[0]
                clsid = row[1]
                data = row[2]
                if clsid in hklm_contents and data != hklm_contents[clsid]:
                    if data != "None":
                        differences.append([user, clsid, data, hklm_contents[clsid]])
        length = len(differences)
                    
    with open(final_csv, "w", newline="") as f:
        writer = csv.writer(f)
        for i in range(length):
            writer.writerow(differences[i])



def main():
    logo = '''
   _____ ____  __  __  _____ _     _      _     _ 
  / ____/ __ \|  \/  |/ ____| |   (_)    | |   | |
 | |   | |  | | \  / | (___ | |__  _  ___| | __| |
 | |   | |  | | |\/| |\___ \| '_ \| |/ _ \ |/ _` |
 | |___| |__| | |  | |____) | | | | |  __/ | (_| |
  \_____\____/|_|  |_|_____/|_| |_|_|\___|_|\__,_|
                                                  
                                                  
 '''

    print(logo)
    init()
    parser = argparse.ArgumentParser(prog='COMShield.py', usage='\t\tpython.exe COMShield.py [options]', description='COMShield tool for Windows compares the registry values of a specific key between HKEY_USERS (HKU) and HKEY_LOCAL_MACHINE (HKLM) in the CLSID registry. It detects any differences between the two and alerts the user if any discrepancies are found. \n\n Emaples:\n\n'
          '\tpython.exe %(prog)s \n'
          '\tpython.exe %(prog)s --path C:\\Users\\COM\\Desktop --output output.csv\n'
          '\tpython.exe %(prog)s --print', formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-v', '--version', action='version', version='%(prog)s 0.1')
    parser.add_argument("--output", help="Name of the CSV file to save the results.",  nargs='?')
    parser.add_argument("--path", help="Path of the CSV file to save the results.")
    parser.add_argument("--print", help="Print the output in the console.", action="store_true")
    args = parser.parse_args()
    
    print(Colors.GREEN + "\t" + "COMShield by Sulaiman Refaee -" + " Version: 0.1" + Colors.ENDC)
    print(Colors.GREEN + "\t" + "Twitter: " + "@sulaimvn" + Colors.ENDC)
    print()
    try:
        if clear_dir:
            shutil.rmtree(clear_dir)
    except OSError:
        pass

    path = os.path.join(os.path.dirname(__file__))

    if " " in path:
        print(ERROR + time.ctime() + " The path has a spaces please runthe program in another path")
        sys.exit(1)
        
    else:
        if not args.output and not args.print:
        
            copy_user_profile_data()

            sids = get_sids()

            for sid in sids:
                run_recmd(sid)
            shutil.rmtree(clear_dir)
            
            get_hklm()
            current_time = datetime.datetime.now()
            formatted_time = current_time.strftime("%Y%m%d_%H%M%S")
            default_name = formatted_time + "_COMShield.csv"
            compare_registry_values(default_name)
            print("Results saved to: " + Colors.GREEN + os.path.join(os.path.dirname(__file__), default_name) + Colors.ENDC)
        
        elif args.print and not args.output:
        
            copy_user_profile_data()

            sids = get_sids()

            for sid in sids:
                run_recmd(sid)
            shutil.rmtree(clear_dir)
            
            get_hklm()
            current_time = datetime.datetime.now()
            formatted_time = current_time.strftime("%Y%m%d_%H%M%S")
            default_name = formatted_time + "_COMShield.csv"
            compare_registry_values(default_name)
            print_file = os.path.join(os.path.dirname(__file__), default_name)
            command = 'type "{}"'.format(print_file)
            subprocess.run(command, shell=True)
            os.remove(print_file)            
        
        else:
            copy_user_profile_data()

            sids = get_sids()

            for sid in sids:
                run_recmd(sid)
            shutil.rmtree(clear_dir)
            
            get_hklm()
            user_name = args.output
            
            compare_registry_values(user_name)
            print("Results saved to: " + Colors.GREEN + os.path.join(os.path.dirname(__file__), args.output) + Colors.ENDC)

        
            
            
            
        hklmfile = os.path.join(os.path.dirname(__file__), "hklm_output.csv")
        allusersfile = os.path.join(os.path.dirname(__file__), "allusers_output.csv")
        os.remove(hklmfile)
        os.remove(allusersfile)
        
        
        
        
try:
    main()
    
except KeyboardInterrupt:
    print(WARNING + time.ctime() + " [CTRL + C]? OK...")
    sys.exit(0)
    
except Exception as e:
    print(ERROR + time.ctime() + " Unknown error: " + str(e))
