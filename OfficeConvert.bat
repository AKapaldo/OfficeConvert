@echo off
:: Office Convert
cls
title Office Convert 1.1
echo Place files in C:\Temp\Office Convert\
echo.
pause
echo.

echo Word - doc to docx conversion running...
for /r "C:\Temp\Office Convert" %%a in ("*.doc") do (
"C:\Program Files (x86)\Microsoft Office\root\Office16\Wordconv.exe" -oics -nme "%%a" "%%ax"
del "%%a"
)
echo Word conversion complete!
echo.

echo Excel - xls to xlsx conversion running...
for /r "C:\Temp\Office Convert" %%a in ("*.xls") do (
"C:\Program Files (x86)\Microsoft Office\root\Office16\excelcnv.exe" -nme -oice "%%a" "%%ax"
del "%%a"
)
echo Excel conversion complete!
echo.
pause
