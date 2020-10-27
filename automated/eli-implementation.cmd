@ECHO OFF
NET SESSION >NUL 
IF %ERRORLEVEL% NEQ 0 GOTO ELEVATE
GOTO ADMINTASKS

:ELEVATE
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
EXIT

:ADMINTASKS
CD "%~dp0"

REM powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%~dp0O2016RTool'"
REM powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%~dp0O2016RTool.zip'"
REM powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%~dp0O2016RTool.zip'"
REM powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%~dp0O2016RTool'"

IF EXIST "C:\Windows\System32\curl.exe" (
	curl -L -O https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/O2016RTool.zip
) ELSE (
	IF EXIST "C:\Windows\System32\bitsadmin.exe" (
		bitsadmin /transfer myDownloadJob /download /priority normal "https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/O2016RTool.zip" "%CD%O2016RTool.zip"
	) ELSE (
		"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --incognito https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/O2016RTool.zip"
	)
)

mkdir "%~dp0O2016RTool"
tar -xf "%~dp0O2016RTool.zip" --directory "%~dp0O2016RTool"


REM powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%~dp0O2016RTool'"
REM powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%~dp0O2016RTool.zip'"


pause
EXIT