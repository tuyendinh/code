echo off
setlocal enabledelayedexpansion enableextensions

REM *****************************************************
REM **************USER SETTING HERE************************
REM Setting variables
set WINAMS_PATH=C:\WinAMS
set PROJECT_PATH=Q:\WinAMS\ISF-ME_3x_RadarSoft
set PROJECT_NAME=ISF-ME_3x_RadarSoft
REM folder contain csv file. The path should not contain module.
REM For example, Q:\RadarSoft_Tuyen_15112017\01_Test_Spec\02_Test_Case\KNL\zynq -> the path should be Q:\RadarSoft_Tuyen_15112017\01_Test_Spec\02_Test_Case
set CSV_PATH=D:\Auto_exeution\Test_case\execute
REM folder contain result of execution
set OUTPUT_PATH=D:\Auto_exeution\Result
set STARTUP_FILE_DEFAULT="Q:\winAMS\RadarSoft\SS_STARTUP.txt"
REM set LOG_FILE="%OUTPUT_PATH%\%DATE%.log"
set AMS_ERROR_LOG="%PROJECT_PATH%\AmsErrorLog.txt"
REM Enter lists of test case you want to run here. Leave it empty if you want to execute all csv file in a folder
set FILE_LIST=
REM or
REM set FILE_LIST=Test_CAN_SetFilter Test_Can_Init_Comm_Timing Test_TS_ExecIdleMode Test_CAN_Reset Test_Can_Init_Comm_Timing

REM **************END************************************

set today=!date:/=!
set now=!time::=!
set LOG_FILE="%OUTPUT_PATH%\Log_!today!_!now!.txt"
set PROJECT="%PROJECT_PATH%\%PROJECT_NAME%.amsy"
set WINAMS_EXE="%WINAMS_PATH%\bin\SSTManager.exe"

REM if FILE_LIST is empty, get file name in the folder c
REM listout all file *.csv in the folder %CSV_PATH%
REM execute test 
set startTime=%time%
REM If the SST is opend
call:KillProcess
set count=1
if "%FILE_LIST%" == "" (
    for /f "delims=" %%f in ('dir /b /s "%CSV_PATH%\*.csv"') do (
        echo !count!. Testcase: %%f
        echo !count!. Testcase: %%f >> %LOG_FILE%
        call :ExecuteTest %%f
		call:KillProcess
		type %AMS_ERROR_LOG% >> %LOG_FILE%
		echo ====================================================================================================================>> %LOG_FILE%
        set /A count+=1
    )
)else (
    for %%f in (%FILE_LIST%) do (
        echo !count!. Testcase: %%f
		echo !count!. Testcase: %%f >> %LOG_FILE%
        set "FilePath="
        call :FindFilePath %%f FilePath
        if "!FilePath!" == "" (
            echo WARNING: The test case %%f does not exist.
            echo WARNING: The test case %%f does not exist.>> %LOG_FILE%
        )else (
            echo Path: !FilePath!
            call :ExecuteTest !FilePath!
        )
		call:KillProcess
		type %AMS_ERROR_LOG% >> %LOG_FILE%
		echo ====================================================================================================================>> %LOG_FILE%
        set /A count+=1
    )
)

echo Finish execution.
set endtime=%time%
REM set /a timetaken=!endtime! - !starttime!
REM echo Time running: !timetaken!
REM Check status of program
:ExecutionTimeout
set TIMEOUT=0
:while
TASKLIST|FINDSTR SSTManager.exe >null
if %ERRORLEVEL%==0 (
REM    timeout /t 1 /NOBREAK >null
REM	ping -n 2 127.0.0.1 > nul
    set /A TIMEOUT+=1
    echo | set /p dummy=.
    if !TIMEOUT!==15000 (
        echo.
        echo ERROR: Terminating execution of current test case because of timeout!
        echo ERROR: Terminating execution of current test case because of timeout! >> %LOG_FILE%
        call:KillProcess
REM        timeout /t 1 /NOBREAK >null
REM		ping -n 2 127.0.0.1 > nul
        call:KillProcess
        goto End
    )
    goto while
)
echo.
goto End

REM Terminate the test cases when timeout.
:KillProcess
TASKLIST|FINDSTR LiX.exe
if %ERRORLEVEL%==0 (
taskkill /IM LiX.exe /f >null
)
TASKLIST|FINDSTR SystemSimulator.exe
if %ERRORLEVEL%==0 (
taskkill /IM SystemSimulator.exe /f >null
title exclude
taskkill /IM cmd.exe /FI "WINDOWTITLE ne exclude*" >null
)
TASKLIST|FINDSTR SSTManager.exe
if %ERRORLEVEL%==0 (
taskkill /IM SSTManager.exe /f >null
)
goto End

REM Procedure for execute test
:ExecuteTest
set tc=%~1
set file=!tc!
REM Get startup file
set StartupFile=!file:.csv=.txt!
if exist !StartupFile! (
    echo The test uses specific startup file: !StartupFile!
)else (
    echo The test use common startup file: !STARTUP_FILE_DEFAULT!
    set StartupFile=!STARTUP_FILE_DEFAULT!
)
REM Create folder to store test result
set file=!file:.csv=!
set file=!!file:%CSV_PATH%\=!!
echo Running with timeout 30s
REM Seting startup file and enable coverage log
%WINAMS_EXE% -set_system_g Start=!StartupFile! -set_test AutoCovLog=1 %PROJECT%
REM %WINAMS_EXE% -set_sampl_system_g Start=!StartupFile! -set_test CoverLogFormat=0 %PROJECT%
%WINAMS_EXE% -set_sampl_system_g Start=!StartupFile! -set_test CoverLogFormat=1 %PROJECT%
REM start "" %WINAMS_EXE% -set_system_g Start=!StartupFile! -b -mcdc -testCsv !tc! -output %OUTPUT_PATH%\!file! %PROJECT%
start "" %WINAMS_EXE% -e -b -mcdc -testCsv !tc! -output %OUTPUT_PATH%\!file! %PROJECT%
call:ExecutionTimeout
goto End

REM Find file path
:FindFilePath
set TcName=%~1
set "p="
for /r %CSV_PATH% %%a in (*.csv) do if /i "%%~nxa"=="!TcName!.csv"  set "p=%%~dpnxa"
if not "!p!" == "" (
    set "%~2=%p%"
)
goto End

REM remove white space
:Trim
SetLocal EnableDelayedExpansion
set Params=%*
for /f "tokens=1*" %%a in ("!Params!") do EndLocal & set %1=%%b
exit /b
:End
