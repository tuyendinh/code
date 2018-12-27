@echo off
setlocal enabledelayedexpansion enableextensions
REM  Use: ExportCSV.bat PathToFileListTestCases FolderReleasePath OutputPath

REM setting 
set UnitTestHelper="C:\Program Files (x86)\UnitTestHelper\UnitTestHelper.exe"

REM IF the path is unchange -> remove below comment and update path
::set TestListFile=F:\Document\winAMS\winAMS_guide\Tools\batch\TestList.txt
::set TestDesignPath=F:\work\3_5G_Soft_2\branches\Phase_3\04_Output
::set OutputDir=C:\Users\tuyendvu\Desktop\export

REM Get parameter via command line if parameter is not set
if "%TestListFile%"=="" (set TestListFile=%1)
if "%TestDesignPath%"=="" (set TestDesignPath=%2)
if "%OutputDir%"=="" (set OutputDir=%3)
IF "%TestDesignPath:~-1%"=="\" SET TestDesignPath=%TestDesignPath:~0,-1%
IF "%OutputDir:~-1%"=="\" SET OutputDir=%OutputDir:~0,-1%
set TestDesignPath=!TestDesignPath!\01_Test_Spec\01_Test_Design
set OutputDir=!OutputDir!\02_TestCases


echo %OutputDir%
echo %TestDesignPath%
set count=0
For /F %%a in (%TestListFile%) DO (
	set /A count+=1
	echo !count! %%a
	set FileName=%%a
	set "TestCasePath="
	call :Trim FileName !FileName!
    call :FindFilePath !TestDesignPath! !FileName! "xlsx" TestCasePath
	
    if "!TestCasePath!" == "" (
        echo ERROR: Can not find test case %%a
    )else (
		Set OutPath=!TestCasePath!
		set OutPath=!OutPath:%TestDesignPath%=%OutputDir%!
		if not exist !OutPath! mkdir !OutPath!
		echo %UnitTestHelper% -e -f=!TestCasePath!!FileName!.xlsx -o=!OutPath! -s
        start "" /B /wait %UnitTestHelper% -e -f=!TestCasePath!!FileName!.xlsx -o=!OutPath!
		REM Checking export sucessfully
		call :FindFilePath !OutPath! !FileName! "csv" TestCasePath
		if !TestCasePath!=="" echo "ERROR: Testcase %%a Not yet complete exporting"
    )
)
goto End

REM Find file path
:FindFilePath
set RootPath=%~1
set TcName=%~2
set ft=%~3
set "p="
for /r %RootPath% %%a in ( dir *.!ft!) do if /i "%%~nxa"=="!TcName!.!ft!" set "p=%%~dpa"
set "%~4=%p%"
goto End
:Trim
SetLocal EnableDelayedExpansion
set Params=%*
for /f "tokens=1*" %%a in ("!Params!") do EndLocal & set %1=%%b
exit /b
:End
