# ESP-Flasher-GUI
ESP device flashing tool (GUI) for production with exel support to save data

### Build for win7 32bit from 64bit machine
Requirements:
- Python 3.8
- wxPython 4.1.1

Python version 3.9 and above is not supported by wxPython 4.1.1. Have dll issues for latest verison of python and wxPython  

Install python 3.8 32bit version from https://www.python.org/downloads/release/python-387/  

Install virtualenv
```
C:\<Path>\Python\Python38-32\Scripts\pip.exe install virtualenv
```
Create virtualenv and activate
```
C:\<Path>\Python\Python38-32\Scripts\virtualenv.exe venv
```
Install requirements
```
pip install -r requirements.txt
```
Install pyinstaller
```
pip install pyinstaller
```
Build
```
pyinstaller .\espflasher.spec
```

Change **pathex=** to project path