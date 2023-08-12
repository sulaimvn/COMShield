
# COMShield
**COMShield** tool that detects differences between HKLM (HKEY_LOCAL_MACHINE) and ALL Users in the registry inside CLSID.

## Overview 
### Description 
**COMShield** tool for Windows compares the registry values of a specific key between HKEY_USERS (HKU) and HKEY_LOCAL_MACHINE (HKLM) in the CLSID registry. It detects any differences between the two and alerts the user if any discrepancies are found.

## Features 
### Current features:
1. Retrieving subkeys from the HKEY_LOCAL_MACHINE\Software\Classes\CLSID and All Users CLSID keys in the Windows registry.
2. Comparing the registry values between the two keys and printing out any differences found.

### Ideas in mind:
1. Detects different persistence attacks in COM Object
2. Monitor the registration and usage of COM Object in Windows Event Log
   
## Preparing
1. Install RawCopy.exe
2. install RECmd.exe
3. install .NET 6
4. <ins>Python 3.11 or newer</ins>. If you're hit with a syntax error on match-case statements, this's why.
5. Using COMShield in a Windows Enviroment.
6. Install required packages:
   ```
   python.exe pip.exe install -r requirements.txt
   ```

## Usage
1. Clone the repository:
   ```
   https://github.com/sulaimvn/COMShield.git
   ```
2. Install the required packages
3. Running COMShield as an **Administrator privileges** 
   ```
   python.exe COMShield.py
   ```
## Contact
Please reach out to me directly about suggestions for **COMShield** at *sulaimvn@gmail.com* or *https://twitter.com/sulaimvn*.
