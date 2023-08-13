
# COMShield
**COMShield** tool that detects differences between HKLM (HKEY_LOCAL_MACHINE) and ALL Users in the registry inside CLSID.

## Overview 
### Description 
**COMShield** is a tool for Windows that compares the registry values of a specific key between HKEY_USERS (HKU) and HKEY_LOCAL_MACHINE (HKLM) in the CLSID registry. It detects any differences between the two which alerts the user if any discrepancies are found.

## Features 
### Current features:
1. Retrieving subkeys from the HKEY_LOCAL_MACHINE\Software\Classes\CLSID and All Users CLSID keys in the Windows registry.
2. Comparing the registry values between the two keys and printing out any differences found.

### Ideas in mind:
1. Detects different persistence attacks in COM Object.
2. Monitor the registration and usage of COM Object in Windows Event Log.
   
## Preparing
1. Install RawCopy.exe
   ```
   Please ensure that the files are downloaded to the directory: C:\*\COMShield\RawCopy-master\ 

   Link:
   https://github.com/jschicht/RawCopy
   ```
2. Install RECmd.exe
   ```
   Please ensure that the files are downloaded to the directory: C:\*\COMShield\RECmd\
   
   Link:
   https://f001.backblazeb2.com/file/EricZimmermanTools/net6/RECmd.zip
   ```
3. Install .NET 6
4. Python 3.11 or newer.
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
2. Running COMShield as an **Administrator privileges** 
   ```
   python.exe COMShield.py
   ```
## Contact
Please reach out to me directly about suggestions for **COMShield** at *sulaimvn[at]gmail[dot]com* or *https://twitter.com/sulaimvn*.
