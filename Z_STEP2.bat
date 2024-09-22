@echo off

:: Set the working directory
::++++++++++work++++++++++++++
:: cd /d "C:\Users\rclausing\DIBBS"
cd /d "C:\Users\aeros\DIBBSBQ"
::++++++++++home++++++++++++++

:: Unzip all *.zip files in the directory
for %%i in (*.zip) do (
    echo Unzipping %%i...
    powershell -command "Expand-Archive -Force '%%i' ." || (
        echo Failed to unzip %%i >> script_log.txt
        exit /b 1
    )
)

:: Combine all bq24*.txt files into a single BQ.TXT file
if exist bq24*.txt (
    copy /b bq24*.txt BQ.TXT
    del /q bq24*.txt
) else (
    echo No bq24*.txt files found. >> script_log.txt
)

:: Delete all as24*.txt files
if exist as24*.txt del /q as24*.txt

:: Delete all BQ*.zip files
if exist BQ*.zip del /q BQ*.zip

:: Kill any running instances of Excel
taskkill /f /im excel.exe

:: Run the next VBScript
::++++++++++++++++++++++++
::cscript //nologo "C:\Users\rclausing\DIBBS\STEPS\Z_STEP3.vbs"
  cscript //nologo "C:\Users\aeros\DIBBSBQ\STEPS\Z_STEP3.vbs"
::++++++++++++++++++++++++
@echo on
