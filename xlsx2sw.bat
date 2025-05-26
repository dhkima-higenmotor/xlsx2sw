@echo off

REM path
set root=%USERPROFILE%\miniforge3
call %root%\Scripts\activate.bat %root%
call conda activate base

call python xlsx2sw.py

pause
