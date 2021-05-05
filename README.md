[![Generic badge](https://img.shields.io/badge/Python-3.9.5-<COLOR>.svg)](https://www.python.org/downloads/release/python-395/)

# Extract PDF from .msg files

Simple python script to extract .pdf attachments from .msg file on Windows

# INSTALLATION

All you need is the win32com module : ```$python -m pip install pywin32```

# HOW IT WORKS 

1. Set the source and destination folder in the script.
2. Launch the script : ```$python extract-pdf-from-msg.py```
3. The script scan the source folder for every .msg file. It then open them and search for .pdf file.
4. The script save every pdf file he found in the source folder to the destination folder and name them like the original .msg file.
