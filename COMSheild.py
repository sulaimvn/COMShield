import winreg
import sys
import time
import argparse
from colorama import init


class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GREY = '\33[90m'
    ENDC = '\033[0m'

WARNING = Colors.YELLOW + "WARNING: " + Colors.ENDC
ALERT = Colors.RED + "ALERT: " + Colors.ENDC
INFO = Colors.GREY + "INFO: " + Colors.ENDC
ERROR = Colors.RED + "ERROR: " + Colors.ENDC


def compare_registry_values():
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
                    i += 1To change the order of the output, you can modify the `compare_registry_values()` function as follows:

```python
def compare_registry_values():
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

    # Sort the dictionaries by keys for consistent output
    sorted_values_local = dict(sorted(values_dict_local.items(), key=lambda x: x[0]))
    sorted_values_current = dict(sorted(values_dict_current.items(), key=lambda x: x[0]))

    # Print the values from HKEY_LOCAL_MACHINE\Software\Classes\CLSID
    print(f"{INFO}Values from HKEY_LOCAL_MACHINE\\Software\\Classes\\CLSID:")
    for key, value in sorted_values_local.items():
        print