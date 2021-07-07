@ECHO OFF
NET SESSION >NUL 
IF %ERRORLEVEL% NEQ 0 GOTO ELEVATE
GOTO ADMINTASKS

:ELEVATE
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
EXIT

:ADMINTASKS
CD "%~dp0"
ECHO %cd%
IF EXIST "C:\Windows\System32\curl.exe" (
	curl -L -O https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip
			powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%CD%\O2016RToolModified.zip'"
) ELSE (
	IF EXIST "C:\Windows\System32\bitsadmin.exe" (
		bitsadmin /transfer myDownloadJob /download /priority normal "https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip" "%CD%O2016RToolModified.zip"
		powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%CD%\O2016RToolModified.zip'"
	) ELSE (
		CD "%USERPROFILE%\Downloads"
		powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%CD%\O2016RToolModified.zip'"
		IF EXIST "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"  --incognito https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip"
		IF EXIST "%ProgramFiles%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe"  --incognito https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip"
		1GB.zip.crdownload
		
	)
)





MKDIR "%cd%\O2016RToolModified"
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%CD%\O2016RToolModified'"
tar -xf %cd%\O2016RToolModified.zip --directory %cd%\
START /WAIT "" ".\O2016RToolModified\O2016RTool.cmd" | BREAK 

RD /S /Q ".\O2016RToolModified"
DEL ".\O2016RToolModified.zip"
powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%CD%\O2016RToolModified'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%CD%\O2016RToolModified.zip'"

REM DEL test.cmd






:waitChromeDownload

IF EXIST "%USERPROFILE%\Downloads\O2016RToolModified.zip.crdownload" (
	ECHO O2016RToolModified.zip has been incompleted by Chrome.
	ECHO Please wait until the download is completed.
	timeout /t 3
	GOTO :waitChromeDownload
)

GOTO :EOF