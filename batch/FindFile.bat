@echo off

:: the file to look for. (In this case 'myFile.txt')
set filename=Test_CAN_Init.txt

:: the drive or path to search. (In this case searching current drive)
set searchPath=D:\04_Work\Mock4\RadarSoft_Tuyen_15112017\02_Test_Result

:: If the file is found. This variable will be set
set foundFilePath=

:: echos all found paths and returns the last occurrance of the file path
FOR /R "%searchPath%" %%a  in (%filename%) DO (
    IF EXIST "%%~fa" (
        echo "%%~fa" 
        SET foundFilePath=%%~fa
    )
)

IF EXIST "%foundFilePath%" (
    echo The foundFilePath var is set to '%foundFilePath%'
) else (
    echo Could not find file '%filename%' under '%searchPath%'
)