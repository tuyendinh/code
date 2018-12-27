echo off

REM =============User setting here===========================
set GHS_DIR=C:\GHS\V800.V2013.5.5\comp_201355
set MCDC_PATH=S:\ISF_ME_2x_D0-2Pre5\MCDC
set SRC_PATH=S:\ISF_ME_2x_D0-2Pre5\SRC

REM =============User end ===========================
set task=%1
REM switch %task
echo task: %task%
if "%task%"=="C0" call:build_C0
if "%task%"=="MCDC" call:build_MCDC
if "%task%"=="all" call:build_C0&call:build_MCDC
if "%task%"=="" call:build_C0&call:build_MCDC
goto end
REM sub routine compile source code for MCDC
:build_C0
echo ========Start compile source code for C0======================================
echo Deleting obj files
del /s/q %SRC_PATH%\out\* 1>null
del /s/q %SRC_PATH%\objs\* 1>null
echo End deleting object file
echo Compiling
%GHS_DIR%\gbuild -top %SRC_PATH%\Apl.gpj > build_src.log 2>&1 & type build_src.log
echo ========End compile source code for C0========================================
goto end 

REM sub routine compile source code for MCDC
:build_MCDC
echo ========Start compile source code for MCDC====================================
echo Deleting obj files
del /s/q %MCDC_PATH%\out\* 1>null
del /s/q %MCDC_PATH%\objs\* 1>null
echo End deleting object file
echo Compiling
%GHS_DIR%\gbuild -top %MCDC_PATH%\Apl.gpj > build_mcdc.log 2>&1 & type build_mcdc.log
echo ========End compile source code for MCDC======================================
goto end 
:end
