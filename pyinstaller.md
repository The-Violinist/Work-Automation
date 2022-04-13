# Syntax for creating an exe

`pyinstaller --onefile -i"\path\of\icon.ico"  path\of\file.py`

If the executable fails due to modules not being included, add the path to the site packages:

`C:\Python310\Lib\site-packages` for example