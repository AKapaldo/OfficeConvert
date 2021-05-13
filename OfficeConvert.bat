@echo off
:: Office Convert
:Start
cls
title Office Convert 1.0
echo Place files in C:\Temp\Office Convert\
echo 
echo 1. Word - doc to docx
echo 2. Excel - xls to xlsx
echo.
CHOICE /C 12 /M "Enter your choice: "
IF ERRORLEVEL 2 GOTO Excel
IF ERRORLEVEL 1 GOTO Word

:Word
cls
title Word - doc to docx
for /r "C:\Temp\Office Convert" %%a in ("*.doc") do (
"C:\Program Files (x86)\Microsoft Office\root\Office16\Wordconv.exe" -oics -nme "%%a" "%%ax"
del "%%a"
)
echo.
echo Conversion complete!
pause
GOTO Start

:Excel
cls
title Excel - xls to xlsx
for /r "C:\Temp\Office Convert" %%a in ("*.xls") do (
"C:\Program Files (x86)\Microsoft Office\root\Office16\excelcnv.exe" -nme -oice "%%a" "%%ax"
del "%%a"
)
echo.
echo Conversion complete!
pause
GOTO Start
