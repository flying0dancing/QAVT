@echo off
REM ===========================================================================
REM             Contact Kun.Shen@lombardrisk.com if any bug    
REM ===========================================================================
title QAVT.bat running...
if not "!PROCESSOR_ARCHITECTURE!"=="%PROCESSOR_ARCHITECTURE%" (
	cmd /V:ON /C %0 %*
    goto EOF
)

:Start


set _QAVT_CONST_CLIENT_SCRIPT_ROOT=%~dp0

set _QAVT_VAR_EXITCODE=
if /I not "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%"=="%SystemDrive%\QAVT_Auto\" (
	echo Error: QAVT.bat must be run from the folder where it is installed!
	echo.
	goto ErrorExit
)

cd /d %_QAVT_CONST_CLIENT_SCRIPT_ROOT%
attrib /s -r *
md QAVT.Output

REM ===========================================================================
REM Reset _QAVT_VAR_*
REM for /f "delims==" %%i in ('set _QAVT_VAR') do (
REM    set %%i=
REM )
REM ===========================================================================


REM echo ##################################
echo Getting latest scripts...
echo Into %_QAVT_CONST_CLIENT_SCRIPT_ROOT%scripts
echo ##################################
IF NOT EXIST %windir%\system32\robocopy.exe ( copy \\sha-sql2005-c\QA\Applications\robocopy.exe %windir%\system32\robocopy.exe /Y 1>nul)
robocopy.exe /TBD /R:30 /W:60 /MIR %_QAVT_CONST_SERVER_SCRIPT_ROOT% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%scripts 1>nul


REM echo ##################################
echo Getting latest INI files...
echo Into %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Config
echo ##################################
set _QAVT_CONST_COMPUTER_NAME=%COMPUTERNAME%
set _QAVT_VAR_CONFIG_FILE=
robocopy.exe /TBD /R:30 /W:60 %_QAVT_CONST_SERVER_CONFIG_ROOT% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Config QAVT.%_QAVT_CONST_COMPUTER_NAME%_*.ini 1>nul

:CONFIG
REM ===========================================================================
REM select current config file
set _QAVT_VAR_NUM_OF_CONFIG=0
set _QAVT_CONST_OUTPUT_EXCEL=
set _QAVT_CONST_OUTPUT_LOG=
set _QAVT_VAR_CONFIG_FILE=
set _QAVT_VAR_CONFIG_NAME=
for /R %_QAVT_CONST_CLIENT_SCRIPT_ROOT%\QAVT.Config %%i in (QAVT.%_QAVT_CONST_COMPUTER_NAME%*.ini) do (
    set /A _QAVT_VAR_NUM_OF_CONFIG+=1
	set _QAVT_VAR_CONFIG_FILE=%%i
) 

echo ##################################
REM echo _QAVT_VAR_NUM_OF_CONFIG is set to %_QAVT_VAR_NUM_OF_CONFIG%

if "%_QAVT_VAR_NUM_OF_CONFIG%"=="0" (

    rem echo Error: _QAVT_VAR_NUM_OF_CONFIG is 0! 
	if exist %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\*.ini (
	move /y %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\*.ini %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Config 1>nul 2>&1
	
	echo.....................................
	echo read all config files.
	echo QAVT.bat -- Finished Running. ^^o^^
	echo.....................................
	goto EOF
	) else (
	echo.....................................
	echo Error: _QAVT_VAR_NUM_OF_CONFIG is set to 0.
	echo Error: _QAVT_VAR_CONFIG_FILE is set to null.
	goto ErrorExit
	)


) ELSE (

echo _QAVT_VAR_CONFIG_FILE is set to %_QAVT_VAR_CONFIG_FILE%

set _QAVT_VAR_CONFIG_NAME=!_QAVT_VAR_CONFIG_FILE:%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Config\QAVT.%_QAVT_CONST_COMPUTER_NAME%_=!
set _QAVT_VAR_CONFIG_NAME=!_QAVT_VAR_CONFIG_NAME:.ini=!

	if /I "!_QAVT_VAR_CONFIG_NAME:~0,11!"=="AnaDecision" (
		move /y %_QAVT_VAR_CONFIG_FILE% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output 1>nul
		set _QAVT_VAR_CONFIG_FILE=%_QAVT_VAR_CONFIG_FILE:\QAVT.Config\=\QAVT.Output\%
		call :AnaDecision
		goto :CONFIG
	) ELSE (
	 if /I "!_QAVT_VAR_CONFIG_NAME:~0,11!"=="CmpDecision" (
		move /y %_QAVT_VAR_CONFIG_FILE% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output 1>nul
		set _QAVT_VAR_CONFIG_FILE=%_QAVT_VAR_CONFIG_FILE:\QAVT.Config\=\QAVT.Output\%
		call :CmpDecision
		goto :CONFIG
	 ) ELSE (
		if /I "!_QAVT_VAR_CONFIG_NAME:~0,11!"=="CmpSumXVals" (
		move /y %_QAVT_VAR_CONFIG_FILE% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output 1>nul
		set _QAVT_VAR_CONFIG_FILE=%_QAVT_VAR_CONFIG_FILE:\QAVT.Config\=\QAVT.Output\%
		call :CmpSumXVals
		goto :CONFIG
	    ) ELSE (
		  if /I "!_QAVT_VAR_CONFIG_NAME:~0,11!"=="ChkTransmit" (
		  move /y %_QAVT_VAR_CONFIG_FILE% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output 1>nul
		  set _QAVT_VAR_CONFIG_FILE=%_QAVT_VAR_CONFIG_FILE:\QAVT.Config\=\QAVT.Output\%
		  call :ChkTransmit
		  goto :CONFIG
	      ) ELSE (
		     move /y %_QAVT_VAR_CONFIG_FILE% %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output 1>nul
		     echo Error: %_QAVT_VAR_CONFIG_FILE% is not belong to which QAVT.bat want to analyze.>>%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_CONST_COMPUTER_NAME%_Error.log
		     echo.....................................
	         echo Error: %_QAVT_VAR_CONFIG_FILE% is not belong to which QAVT.bat want to analyze.
		     echo    other config files will be continued executed, 
		     ECHO       please check QAVT.%_QAVT_CONST_COMPUTER_NAME%_Error.log.
		     goto :CONFIG
	      )
	    )
	 )
	)
)

GOTO :EOF


:AnaDecision
rem echo ##################################
echo Calling AnaDecision scripts...
echo Restore results at %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output
rem echo ##################################
rem FOR /F "tokens=1,2,* delims== " %%i in ('type %_QAVT_VAR_CONFIG_FILE%') do (
rem  set %%i=%%j
rem  echo %%i is set to %%j
rem )

call :SetPaths
rem echo ##################################
echo _QAVT_CONST_OUTPUT_EXCEL is set to %_QAVT_CONST_OUTPUT_EXCEL%
echo _QAVT_CONST_OUTPUT_LOG is set to %_QAVT_CONST_OUTPUT_LOG%
rem echo ##################################

 cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\Pre_AnaDecision.vbs" %_QAVT_CONST_OUTPUT_EXCEL%  %_QAVT_CONST_OUTPUT_LOG% %_QAVT_VAR_CONFIG_FILE%
 rem set _QAVT_VAR_COUNT_OF_ROWS=%errorlevel%
 if %errorlevel% GTR 0 (
	cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\AnaDecision.vbs" %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_CONST_OUTPUT_LOG% %_QAVT_VAR_CONFIG_FILE%
  )

GOTO :EOF

:CmpDecision
REM echo ##################################
echo Calling CmpDecision scripts...
echo Restore results at %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output
rem echo ##################################

call :SetPaths
rem echo ##################################
echo _QAVT_CONST_OUTPUT_EXCEL is set to %_QAVT_CONST_OUTPUT_EXCEL%
echo _QAVT_CONST_OUTPUT_LOG is set to %_QAVT_CONST_OUTPUT_LOG%
rem echo ##################################

echo ******This information for QAVT Compared Decision table****** >%_QAVT_CONST_OUTPUT_LOG%
cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\CmpDecision.vbs" %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_VAR_CONFIG_FILE% >>%_QAVT_CONST_OUTPUT_LOG%
rem FOR /F "delims=" %%i in ('cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\CmpDecision.vbs" %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_CONFIG_DBSERVER_INSTANCE% %_QAVT_CONFIG_CMPDB% %_QAVT_CONFIG_DATABASE% %_QAVT_CONFIG_FORM% %_QAVT_CONFIG_DBSERVER_USER% %_QAVT_CONFIG_DBSERVER_PASSWORD%') do (
rem  echo %%i
rem  echo %%i >>%_QAVT_CONST_OUTPUT_LOG%
rem )

GOTO :EOF
:CmpSumXVals
REM echo ##################################
echo Calling CmpSumXVals scripts...
echo Restore results at %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output
rem echo ##################################
FOR /F "tokens=1,2,* delims== " %%i in ('type %_QAVT_VAR_CONFIG_FILE%') do (
  set %%i=%%j
  echo %%i is set to %%j
)

call :SetPaths
rem echo ##################################
echo _QAVT_CONST_OUTPUT_EXCEL is set to %_QAVT_CONST_OUTPUT_EXCEL%
echo _QAVT_CONST_OUTPUT_LOG is set to %_QAVT_CONST_OUTPUT_LOG%
rem echo ##################################

echo  This information for QAVT Compared Sum or validation or cross-validation rules >%_QAVT_CONST_OUTPUT_LOG%
 rem cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\CmpSumXVals.vbs" %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_CONFIG_DBSERVER_INSTANCE% %_QAVT_CONFIG_CMPDB% %_QAVT_CONFIG_DATABASE% %_QAVT_CONFIG_TABLE% %_QAVT_CONFIG_DBSERVER_USER% %_QAVT_CONFIG_DBSERVER_PASSWORD% 
FOR /F "delims=" %%i in ('cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\CmpSumXVals.vbs" %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_CONFIG_DBSERVER_INSTANCE% %_QAVT_CONFIG_CMPDB% %_QAVT_CONFIG_DATABASE% %_QAVT_CONFIG_FORM% %_QAVT_CONFIG_DBSERVER_USER% %_QAVT_CONFIG_DBSERVER_PASSWORD%') do (
  echo %%i
  echo %%i >>%_QAVT_CONST_OUTPUT_LOG%
)
GOTO :EOF
:ChkTransmit
REM echo ##################################
echo Calling ChkTransmit scripts...
echo Restore results at %_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output
rem echo ##################################
FOR /F "tokens=1,2,* delims==" %%i in ('type %_QAVT_VAR_CONFIG_FILE%') do (
  set %%i=%%j
  echo %%i is set to %%j
)

call :SetPaths
rem echo ##################################
set _QAVT_CONST_OUTPUT_EXCEL=%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_VAR_CONFIG_NAME:~0,11%_%_QAVT_CONFIG_FORM%_%_QAVT_CONFIG_TRANSMIT_SHEET%_output_%__DAY%.xls
echo _QAVT_CONST_OUTPUT_EXCEL is set to %_QAVT_CONST_OUTPUT_EXCEL%
echo _QAVT_CONST_OUTPUT_LOG is set to %_QAVT_CONST_OUTPUT_LOG%
rem echo ##################################

cscript //nologo "%_QAVT_CONST_CLIENT_SCRIPT_ROOT%\scripts\ChkTransmit.vbs" %_QAVT_CONFIG_TRANSMIT_TEMPLATE_EXCEL% %_QAVT_CONFIG_TRANSMIT_SHEET% %_QAVT_CONFIG_DBSERVER_INSTANCE% %_QAVT_CONFIG_DATABASE% %_QAVT_CONFIG_FORM% %_QAVT_CONFIG_TRANSMIT_EXCEL% %_QAVT_CONST_OUTPUT_EXCEL% %_QAVT_CONST_OUTPUT_LOG% %_QAVT_CONFIG_DBSERVER_USER% %_QAVT_CONFIG_DBSERVER_PASSWORD%

GOTO :EOF
:SetPaths
FOR /F "tokens=1,2,3* delims=-/ " %%i in ('ECHO %DATE%') do (
	set __DAY=%%i%%j%%k
)
FOR /F "tokens=1,2,3* delims=:. " %%i in ('ECHO %TIME%') do (
	set __TIME=%%i%%j%%k
)

rem set _QAVT_CONST_OUTPUT_EXCEL=%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_VAR_CONFIG_NAME:~0,11%_output_%__DAY%_%__TIME%.xlsX
rem set _QAVT_CONST_OUTPUT_LOG=%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_VAR_CONFIG_NAME:~0,11%_%_QAVT_CONFIG_FORM%_output_%__DAY%_%__TIME%.log
set _QAVT_CONST_OUTPUT_EXCEL=%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_VAR_CONFIG_NAME%_output_%__DAY%.xlsx
set _QAVT_CONST_OUTPUT_LOG=%_QAVT_CONST_CLIENT_SCRIPT_ROOT%QAVT.Output\QAVT.%_QAVT_VAR_CONFIG_NAME%_%__DAY%.log
GOTO :EOF

:ErrorExit
echo.
echo QAVT.bat -- Error encountered!  Exiting...
set _QAVT_VAR_EXITCODE=666
GOTO EOF

:EOF