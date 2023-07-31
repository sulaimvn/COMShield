<picture>
  <img alt="Logo" src="/media/COMShield.jpg" width="900" height="400">
</picture>  

# COMShield
**COMShield** tool that detects differences between HKLM (HKEY_LOCAL_MACHINE) and HKCU (HKEY_CURRENT_USER) in the registry inside CLSID.

## Overview 
### Description 
**COMShield** detects differences between the HKEY_LOCAL_MACHINE and HKEY_CURRENT_USER registry hives. Specifically, it looks for differences in the CLSID subkey of the Software\Classes key in both hives.

## Features 
### Current features:
1. Retrieving subkeys from the HKEY_LOCAL_MACHINE\Software\Classes\CLSID and HKEY_CURRENT_USER\Software\Classes\CLSID keys in the Windows registry.
2. Comparing the registry values between the two keys and printing out any differences found.

### Ideas in mind:
1. Detects different persistence attacks in COM Object
2. Monitor the registration and usage of COM Object in Windows Event Log

## Requirements
1. <ins>Python 3.11 or newer</ins>. If you're hit with a syntax error on match-case statements, this's why.
2. Using COMShield in a Windows Enviroment. 
3. Install required packages:
   ```
   python.exe pip.exe install -r requirements.txt
   ```

## Usage
1. Clone the repository:
   ```
   https://github.com/sulaimvn/COMShield.git
   ```
2. Install the required packages
3. Running COMShield in **Administrator privileges** 
   ```
   python.exe COMShield.py
   ```
## Contact
Please reach out to me directly about suggestions for **COMShield** at *sulaimvn@gmail.com* or *https://twitter.com/sulaimvn*.
