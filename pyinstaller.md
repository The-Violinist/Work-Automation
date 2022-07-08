### General syntax for using pyinstaller to create an exe

- `pyinstaller --onefile -i"\path\of\icon.ico"  path\of\file.py`

### If the executable fails due to modules not being included, add the path to the site packages:

- add `C:\Python310\Lib\site-packages` for example to Analysis->pathex in the .spec file, then rerun pyinstaller with the following command:

  -`pyinstaller file_name.spec`

### If pyinstaller is failing while running, try the following:
- Find the path to the site package: `python -m site`
- Navigate in the `site-package` directory to `pywin32_system32`
  - copy `pythoncom39.dll` and `pywintypes39.dll`
- Navigate up a level and paste the files in `win32`
