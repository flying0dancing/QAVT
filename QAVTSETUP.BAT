REM . do not modify this file.
REM set _QAVT_CONST_SERVER_CONFIG_ROOT=\\sha-sql2005-c\QA\CN\CNBAK\QAVTConfig
REM set _QAVT_CONST_SERVER_SCRIPT_ROOT=\\sha-sql2005-c\QA\CN\CNBAK\scripts\common
set _QAVT_CONST_SERVER_CONFIG_ROOT=D:\WorkArea\CN_Report\QAVT\Config
set _QAVT_CONST_SERVER_SCRIPT_ROOT=D:\WorkArea\CN_Report\QAVT\scripts

IF NOT EXIST %SystemDrive%\QAVT_Auto (
    md %SystemDrive%\QAVT_Auto
) ELSE (
    DEL /F /S /Q %SystemDrive%\QAVT_Auto 1>nul
	RMDIR /S /Q %SystemDrive%\QAVT_Auto 1>nul
	md %SystemDrive%\QAVT_Auto
)

copy %_QAVT_CONST_SERVER_SCRIPT_ROOT%\QAVT.bat %SystemDrive%\QAVT_Auto\QAVT.bat /Y
PUSHD %SystemDrive%\QAVT_Auto
start /min /low cmd /k %SystemDrive%\QAVT_Auto\QAVT.bat
