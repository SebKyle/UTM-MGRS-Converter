# Converter for UTM to MGRS, MGRS to UTM with Excel Integration for AFATDS

This repository is specifically designed for importing and exporting graphics with the AFATDS (Advanced Field Artillery Tactical Data System) in the Army Mission Command System. It uses a Python script to convert between UTM (Universal Transverse Mercator) and MGRS (Military Grid Reference System) coordinates.

## Dependencies:
- Python
- Python MGRS module
- Python OpenPyXL module

## How to Install Dependencies:

1. **Ensure you’re using the latest version of Python**:
   ```bash
   pip install --upgrade pip


2. **If you do not have python installed**:
Be sure to install it from the Python official Website if on windows.
- Linux(Unbuntu/Debian based):
  ```bash
  sudo apt update & sudo apt install python3
- Linux(Fedora-based, using dnf):
  ```bash
  sudo dnf install python3
- Linux(Arch-based using Pacman):
  ```bash
  sudo pacman -S python3
- MacOS:
  ```bash
  brew install python

3. **Install openpyxl**:
   ```bash
   pip install openpyxl 

4. **Install MGRS module**:
   ```bash
     pip install mgrs
