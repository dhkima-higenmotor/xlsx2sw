pyinstaller --onefile --noconsole xlsx2sw.py --hidden-import openpyxl.cell._writer
COPY /Y .\dist\xlsx2sw.exe .\
RMDIR /S /Q dist
RMDIR /S /Q build
DEL /Q xlsx2sw.spec
