REM pyinstaller --onefile --noconsole --upx-dir=%userprofile%\scoop\apps\upx\current xlsx2sw.py --hidden-import openpyxl.cell._writer
pyinstaller --noconsole --upx-dir=%userprofile%\scoop\apps\upx\current xlsx2sw.py --hidden-import openpyxl.cell._writer
sleep 1s
REM COPY /Y .\dist\xlsx2sw.exe .\
REM RMDIR /S /Q dist
REM RMDIR /S /Q build
REM DEL /Q xlsx2sw.spec
