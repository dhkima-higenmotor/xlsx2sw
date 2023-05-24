@echo off

call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat
call python xlsx2sw.py

pause
