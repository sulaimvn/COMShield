import winreg
import sys
import time
import argparse
from colorama import init
import csv
import os
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GREY = '\033[90m'
    ENDC = '\033[0m'
    
WARNING = Colors.YELLOW + "WARNING: " + Colors.ENDC
ALERT = Colors.RED + "ALERT: " + Colors.ENDC
INFO = Colors.GREY + "INFO: " + Colors.ENDC
ERROR = Colors.RED + "ERROR: " + Colors.ENDC

def time_now():
    return time.strftime("[%Y-%m-%d %H:%M:%S] ")

def compare_registry_values(csv_path=None, csv_filename=None):
    # Initialize a list to store the output data
    csv_rows = []
    time.sleep(1)
    # Get the subkeys in HKEY_LOCAL_MACHINE\Software\Classes\CLSID
    key_path_local = r"Software\Classes\CLSID"
    key_local = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path_local, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
    subkey_names_local = []
    i = 0
    
    while True:
        try:
            subkey_name_local = winreg.EnumKey(key_local, i)
            subkey_names_local.append(subkey_name_local)
        except WindowsError as e:
            if e.winerror == 259:  # No more data is available
                break
        i += 1
    winreg.CloseKey(key_local)

    # Get the subkeys in HKEY_CURRENT_USER\Software\Classes\CLSID
    key_path_current = r"Software\Classes\CLSID"
    key_current = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path_current, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
    subkey_names_current = []
    i = 0
    while True:
        try:
            subkey_name_current = winreg.EnumKey(key_current, i)
            subkey_names_current.append(subkey_name_current)
        except WindowsError as e:
            if e.winerror == 259:  # No more data is available
                break
        i += 1
    winreg.CloseKey(key_current)

    # Find the subkeys that are present in both HKEY_LOCAL_MACHINE and HKEY_CURRENT_USER
    common_subkeys = set(subkey_names_local).intersection(set(subkey_names_current))

    # Store the values from HKEY_LOCAL_MACHINE\Software\Classes\CLSID in a dictionary
    values_dict_local = {}
    for subkey_name in common_subkeys:
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
            print(f"Error accessing key {key_path_local}: {e}")
        finally:
            winreg.CloseKey(key_local)

    # Store the values from HKEY_CURRENT_USER\Software\Classes\CLSID in a dictionary
    values_dict_current = {}
    for subkey_name_current in common_subkeys:
        key_path_current = fr"Software\Classes\CLSID\{subkey_name_current}\InProcServer32"
        try:
            key_current = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path_current, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        except FileNotFoundError as e:
            continue
        except PermissionError as e:
            continue

        try:
            num_values = winreg.QueryInfoKey(key_current)[1]
            if num_values > 0:
                i = 0
                while True:
                    try:
                        value_name, value_data, value_type = winreg.EnumValue(key_current, i)
                        if value_name == "":
                            value_name = "Default"
                        values_dict_current[f"{subkey_name_current}_{value_name}"] = value_data
                    except WindowsError as e:
                        if e.winerror == 259:
                            break
                    i += 1
        except Exception as e:
            print(f"Error accessing key {key_path_current}: {e}")
        finally:
            winreg.CloseKey(key_current)

    # Compare the values in the two dictionaries and print out any differences found
    diff_dict = {}
    for key, value in values_dict_local.items():
        if key in values_dict_current and value != values_dict_current[key]:
            diff_dict[key] = (value, values_dict_current[key])
        elif key not in values_dict_current:
            diff_dict[key] = (value, None)

    for key, value in values_dict_current.items():
        if key not in values_dict_local:
           diff_dict[key] = (None, value)

    if len(diff_dict) > 0:
        csv_rows.append(["HKLM Path" , "HKLM default" , "HKCU Path", "HKCU default"])
        for key, values in diff_dict.items():
            if "_" in key and key.split("_")[1]:
                subkey = key.split("_")[0]
                value_name = key.split("_")[1]
                subkey_path = fr"Software\Classes\CLSID\{subkey}\InProcServer32"
            else:
                subkey = key
                value_name = "Default"
                subkey_path = fr"Software\Classes\CLSID\{subkey}\InProcServer32"

            local_machine_default = values[0] if values[0] is not None else "None"
            curr_user_default = values[1] if values[1] is not None else "None"

            if local_machine_default != curr_user_default and local_machine_default not in ["Apartment", "Both", None] and curr_user_default not in ["Apartment","Both", None]:
                if csv_path and csv_filename:
                    csv_rows.append([f"HKEY_LOCAL_MACHINE\\{subkey_path}", local_machine_default, f"HKEY_CURRENT_USER\\{subkey_path}", curr_user_default])
                else:
                    print()
                    time.sleep(1)
                    print(ALERT + f"\tIN HKEY_LOCAL_MACHINE\\{subkey_path}: {Colors.RED + local_machine_default + Colors.ENDC}")
                    time.sleep(1)
                    print(ALERT + f"\tIN HKEY_CURRENT_USER\\{subkey_path}: {Colors.RED + curr_user_default + Colors.ENDC}")

    else:
        csv_rows.append(["No registry value differences found."])
        print(WARNING + time_now() + "No registry value differences found.")
        
    # Write the output data to a CSV file
    if csv_path and csv_filename:
        csv_file = os.path.join(csv_path, csv_filename)
        with open(csv_file, mode='w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(csv_rows)
        
            
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
    parser = argparse.ArgumentParser(prog='COMShield', usage='python.exe COMShield.csv -HKCU --csv "C:\\temp" --csvf COMShield.csv', description='COMShield tool for Windows compares the registry values of a specific key between HKEY_LOCAL_MACHINE (HKLM) and HKEY_CURRENT_USER (HKCU) in the CLSID registry. It detects any differences between the two and alerts the user if any discrepancies are found.')
    parser.add_argument('-v', '--version', action='version', version='%(prog)s 0.1')
    parser.add_argument('--HKCU', help='Compare registry values in HKEY_CURRENT_USER with HKEY_LOCAL_MACHINE', action='store_true')
    parser.add_argument("--csv", help="Path to the CSV file to save the results")
    parser.add_argument("--csvf", help="Name of the CSV file to save the results") 
    parser.add_argument("-p",  help="Print the results to the console", action='store_true')
    args = parser.parse_args()

    init()
    print(Colors.GREEN + "\t\t" + "Sulaiman - Haboob Team" + "  Version: 0.1" + Colors.ENDC)
    print()
    
    if args.HKCU and args.csv and args.csvf:
        compare_registry_values(csv_path=args.csv, csv_filename=args.csvf)
        print("Results saved to: "+ Colors.GREEN +  args.csv +"\\" +args.csvf + Colors.ENDC)
    elif args.HKCU and args.p:
        compare_registry_values(csv_path=None, csv_filename=None)

    else:
        print("For help: python.exe COMShield.exe --help")

try:
    main()
except KeyboardInterrupt:
    print(WARNING + time.ctime() + " [CTRL + C]? OK...")
    sys.exit(0)
except Exception as e:
    print(ERROR + time.ctime() + " Unknown error: " + str(e))
