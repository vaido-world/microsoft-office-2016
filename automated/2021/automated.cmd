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
	curl -L -O https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip
			powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%cd%\O2016RToolModified.zip'"
) ELSE (
	IF EXIST "%SystemRoot%\System32\bitsadmin.exe" (
		powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%cd%\O2016RToolModified.zip'"

		bitsadmin /transfer myDownloadJob /download /priority normal "https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip" "%cd%\O2016RToolModified.zip"
		WHERE tar >NUL 2>NUL
		IF ERRORLEVEL == 1 (
			ECHO Tar Archiver is not available 
			bitsadmin /transfer myDownloadJob /download /priority normal "https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/tar.cab" "%cd%\tar.cab"
			MKDIR "%cd%\tar"
			expand "tar.cab" -F:* "%cd%\tar"
			SET "PATH=%PATH%;%cd%\tar"
			
		) ELSE (
			ECHO Tar Archiver is available.
		)
		
		
		
	) ELSE (
		CD "%USERPROFILE%\Downloads"
		powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%cd%\O2016RToolModified.zip'"
		IF EXIST "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"  --incognito https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip"
		IF EXIST "%ProgramFiles%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe"  --incognito https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/O2016RToolModified.zip"
		IF EXIST "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/tar.cab"
		IF EXIST "%ProgramFiles%\Google\Chrome\Application\chrome.exe" START /MIN /WAIT "" "https://github.com/vaido-world/microsoft-office-2016/raw/BoQsc-patch-1/automated/2021/tar.cab"
		CALL :waitChromeDownload
		
		MKDIR "%cd%\tar"
		expand "tar.cab" -F:* "%cd%\tar"
		SET "PATH=%PATH%;%cd%\tar"
	)
)



MKDIR "%cd%\O2016RToolModified"
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%cd%\O2016RToolModified'"
tar -xf %cd%\O2016RToolModified.zip --directory %cd%\
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%SystemRoot%\System32\SppExtComObjHook.dll'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath '%SystemRoot%\System32\SppExtComObjPatcher.exe'"
START /WAIT "" ".\O2016RToolModified\O2016RTool.cmd" | BREAK 

RD /S /Q ".\O2016RToolModified"
DEL ".\O2016RToolModified.zip"

powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%cd%\O2016RToolModified'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath '%cd%\O2016RToolModified.zip'"

RD /S /Q "%USERPROFILE%\Downloads\tar"
DEL "%USERPROFILE%\Downloads\tar.cab"
DEL "%~dp0%~n0%~x0"





EXIT
:waitChromeDownload

IF EXIST "%USERPROFILE%\Downloads\O2016RToolModified.zip.crdownload" (
	ECHO O2016RToolModified.zip has been incompleted by Chrome.
	ECHO Please wait until the O2016RToolModified.zip download is completed.
	timeout /t 3
	GOTO :waitChromeDownload
)
IF EXIST "%USERPROFILE%\Downloads\tar.cab.crdownload" (
	ECHO O2016RToolModified.zip has been incompleted by Chrome.
	ECHO Please wait until the tar.cab download is completed.
	timeout /t 3
	GOTO :waitChromeDownload
)

GOTO :EOF
