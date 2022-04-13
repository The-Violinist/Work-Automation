# Syntax for creating an exe

`pyinstaller --onefile -i"\path\of\icon.ico"  path\of\file.py`

If the executable fails due to modules not being included, add the path to the site packages:

add `C:\Python310\Lib\site-packages` for example

to Analysis->pathex in the .spec file, then rerun pyinstaller with the following command:

`pyinstaller file_name.spec`
