@echo off
setlocal enabledelayedexpansion

:: Get the current date in yymmdd format
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set year=!datetime:~0,4!
set month=!datetime:~4,2!
set day=!datetime:~6,2!

set formatted_date=%year%%month%%day%

:: Set the source and destination file paths
set source_file=C:\Users\aeros\Apps\DIBBS\LQ.txt
set destination_dir=C:\Users\aeros\Apps\DIBBS\STEPS\LQ
set destination_file=LQ_%formatted_date%.txt

:: Check if the source file exists
if exist "%source_file%" (
    echo Source file exists: %source_file%
    
    :: Rename the file
    ren "%source_file%" "%destination_file%"
    
    :: Move the renamed file to the destination directory
    move "C:\Users\aeros\DIBBS\%destination_file%" "%destination_dir%"
    
    if exist "%destination_dir%\%destination_file%" (
        echo File successfully moved to %destination_dir%\%destination_file%
    )
)

:: Run the Python script Z_STEP5.py
python "C:\Users\aeros\DIBBS\STEPS\Z_STEP5.py"

endlocal
