:================================================================================================================
::===============================================================================================================
:: GET ADMIN RIGHTS
	@echo off
	cls
	set "o2016rtoolpath=%~dp0"
	set "o2016rtoolpath=%o2016rtoolpath:~0,-1%"
	set "o2016rtoolname=%~n0.cmd"
	set "pswindowtitle=$Host.UI.RawUI.WindowTitle = 'Administrator: O2016RTool - 2015/DEC/20'"
::===============================================================================================================
	"%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system" >nul 2>&1
	if "%errorlevel%" NEQ "0" (
	echo:
	echo   Requesting Administrative Privileges...
	echo   Press YES in UAC Prompt to Continue
	echo:
    goto:UACPrompt
	) else (goto:gotAdmin)
::===============================================================================================================
:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%o2016rtoolname%", "", "%o2016rtoolpath%", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B
::===============================================================================================================
:gotAdmin
    if exist "%temp%\getadmin.vbs" (del "%temp%\getadmin.vbs")
	setLocal EnableExtensions EnableDelayedExpansion
	pushd "%o2016rtoolpath%"
	cls
	mode con cols=80 lines=47
	color 1F
:================================================================================================================
::===============================================================================================================
:: DEFINE SYSTEM ENVIRONMENT
	for /f "tokens=6 delims=[]. " %%A in ('ver') do set /a win=%%A
	if %win% LSS 7601 ((echo:)&&(echo:)&&(echo Unsupported Windows detected)&&(echo:)&&(echo Office 2016 minimum OS must be Windows 7 SP1 or better)&&(echo:)&&(goto:TheEndIsNear))
	for /f "tokens=2 delims==" %%a in ('wmic path Win32_Processor get AddressWidth /value') do (set winx=win_x%%a)
	set "sls=SoftwareLicensingService"
	set "slp=SoftwareLicensingProduct"
	set "osps=OfficeSoftwareProtectionService"
	set "ospp=OfficeSoftwareProtectionProduct"
	set "slsversion="
	set "ospsversion="
	for /f "tokens=2 delims==" %%A IN ('"wmic path %sls% get version /format:list"') do set slsversion=%%A
	if %win% LSS 9200 (for /f "tokens=2 delims==" %%A IN ('"wmic path %osps% get version /format:list"') do set ospsversion=%%A)
	set "downpath=not set"
    set "o16insbuild=16.0.6366.2062"
	set "o16curbuild=16.0.6366.2062"
	set "o16build=not set"
	set "o16lang=en-US"
    set "o16lcid=1033"
	set "o16updlocid=not set"
:================================================================================================================
::===============================================================================================================
:Office16VnextInstall
	pushd %o2016rtoolpath%
	mode con cols=80 lines=46
	color 1F
	title O2016RTool - 2016/FEB/01
	cls
    echo:
    echo ============= OFFICE 2016 DOWNLOAD AND INSTALL ===================
    echo __________________________________________________________________
	echo:
	echo [D] DOWNLOAD VNEXT OFFICE OFFLINE INSTALL
    echo:
	echo [I] INSTALL VNEXT OFFICE
	echo:
	echo [C] CONVERT VNEXT AND WEB OFFICE TO VOLUME
	echo:
	echo [K] START KMS ACTIVATION
    echo __________________________________________________________________
	echo:
    echo [O] ONLINE WEB-OFFICE INSTALLER LINK CREATE
    echo __________________________________________________________________
	echo:
    echo [E] END - STOP AND LEAVE PROGRAM
	echo:
    echo ==================================================================
	echo:
	goto :StartKMSActivation
    CHOICE /C DICKOE /N /M "YOUR CHOICE ?"
    if %errorlevel%==1 goto :DownloadO16Offline
	if %errorlevel%==2 goto :InstallO16
    if %errorlevel%==3 goto :Convert16Activate
	if %errorlevel%==4 goto :StartKMSActivation
    if %errorlevel%==5 goto :DownloadO16Online
    if %errorlevel%==6 goto :TheEndIsNear
:================================================================================================================
::===============================================================================================================
:DownloadO16Offline
    pushd %o2016rtoolpath%
	set "o16arch=x86_x64"
    set "installtrigger=X"
	set "branchtrigger=X"
	cls
	echo:
    echo ============ DOWNLOAD OFFICE 16 OFFLINE SETUP ====================
    echo __________________________________________________________________
	echo:
	echo Download Path:  %downpath%
    echo:
	if "%o16updlocid%"=="64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Branch ID:      %o16updlocid% (=Insider/Preview) && goto:DownOfflineContinue
	if "%o16updlocid%"=="492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Branch ID:      %o16updlocid% (=Current/Retail) && goto:DownOfflineContinue
	if "%o16updlocid%"=="not set" echo Branch ID:      %o16updlocid% && goto:DownOfflineContinue
	echo Branch ID:      %o16updlocid% (=Manual Override)
 :DownOfflineContinue
	echo:
	echo Office build:   %o16build%
	echo:
	echo Language:       %o16lang%
    echo:
	echo Architecture:   %o16arch%
    echo __________________________________________________________________
	echo:
	set /p downpath=Set Download Path ^>
	set "downpath=%downpath:"=%"
	set "downdrive=%downpath:~0,2%"
	cd /d %downdrive% >nul 2>&1
	if errorlevel 1 (echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:DownloadO16Offline)
	if "%downpath:~-1%"=="\" set "downpath=%downpath:~0,-1%"
	pushd %o2016rtoolpath%
	echo:
	set /p branchtrigger=Set Branch ID (I = Insider/Preview, C = Current/Retail or M = manual) ^>
	if "%branchtrigger%"=="I" (set "o16updlocid=64256afe-f5d9-4f86-8936-8840a6a4f5be")&&(set "o16build=%o16insbuild%")&&(goto:BranchSelected)
	if "%branchtrigger%"=="i" (set "o16updlocid=64256afe-f5d9-4f86-8936-8840a6a4f5be")&&(set "o16build=%o16insbuild%")&&(goto:BranchSelected)
	if "%branchtrigger%"=="C" (set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60")&&(set "o16build=%o16curbuild%")&&(goto:BranchSelected)
	if "%branchtrigger%"=="c" (set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60")&&(set "o16build=%o16curbuild%")&&(goto:BranchSelected)
	if "%branchtrigger%"=="M" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:ManualBranchSelect)
	if "%branchtrigger%"=="m" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:ManualBranchSelect)
	goto:DownloadO16Offline
:ManualBranchSelect
    echo:
	set /p o16updlocid=Set Branch ID (manual overrride) ^>
	if "%o16updlocid%"=="not set" goto:DownloadO16Offline
:BranchSelected
	set "o16downloadloc=officecdn.microsoft.com/pr/%o16updlocid%/Office/Data"
	echo:
    set /p o16build=Set Office Build ^>
	if "%o16build%"=="not set" goto:DownloadO16Offline
    echo:
::===============================================================================================================
	call :seto16language
::===============================================================================================================
    echo __________________________________________________________________
    echo:
    echo ============ Pending Download (SUMMARY) ==========================
    echo:
    echo Download Path:  %downpath%
    echo:
	if "%o16updlocid%"=="64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Branch ID:      %o16updlocid% (=Insider/Preview) && goto:PendDownContinue
	if "%o16updlocid%"=="492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Branch ID:      %o16updlocid% (=Current/Retail) && goto:PendDownContinue
	if "%o16updlocid%"=="not set" echo Branch ID:      %o16updlocid% && goto:PendDownContinue
	echo Branch ID:      %o16updlocid% (=Manual/Override)
 :PendDownContinue
	echo:
	echo Office Build:   %o16build%
	echo:
	echo Language:       %o16lang%
    echo:
	echo Architecture:   %o16arch%
    echo __________________________________________________________________
	echo:
    set /p installtrigger=Start download now? (1/0) ^>
    if "%installtrigger%"=="0" goto:Office16VnextInstall
    if "%installtrigger%"=="1" goto:Office16VNextDownload
    goto:DownloadO16Offline
:================================================================================================================
::===============================================================================================================
:Office16VNextDownload
	cls
	echo:
	echo ============ DOWNLOADING OFFICE 16 OFFLINE SETUP =================
	echo __________________________________________________________________
	if "%o16updlocid%"=="64256afe-f5d9-4f86-8936-8840a6a4f5be" set "downbranch=InsiderBranch" && goto:ContVNextDownload
	if "%o16updlocid%"=="492350f6-3a01-4f97-b9c0-c7c6ddf67d60" set "downbranch=CurrentBranch" && goto:ContVNextDownload
	set "downbranch=ManualBranch"
 :ContVNextDownload
	md "%downpath%" >nul 2>&1
	cd /d "%downpath%" >nul 2>&1
	mode con cols=147
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/c2rfireflydata.xml
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/c2rfireflydata.xml"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/v32.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v32.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/v32_%o16build%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v32_%o16build%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/v64.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v64.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/v64_%o16build%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v64_%o16build%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/fre.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/fre.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/i320.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/i320.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/i32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/i32%o16lcid%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/i640.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/i640.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/i64%o16lcid%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/i64%o16lcid%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/s320.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/s320.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/s32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/s32%o16lcid%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/s640.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/s640.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/s64%o16lcid%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/s64%o16lcid%.cab"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x64.%o16lang%.dat
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x64.%o16lang%.dat"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat.cobra
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x64.x-none.dat.cobra"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x86.%o16lang%.dat
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x86.%o16lang%.dat"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat"
	"%o2016rtoolpath%\OfficeFixes\wget.exe" --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%o16build%_%o16lang%_x86_x64_%downbranch% http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat.cobra
    if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%o16build%/stream.x86.x-none.dat.cobra"
	echo __________________________________________________________________
	echo %o16build%>%o16build%_%o16lang%_x86_x64_%downbranch%\package.info
	echo %o16lang%>>%o16build%_%o16lang%_x86_x64_%downbranch%\package.info
	echo %o16updlocid%>>%o16build%_%o16lang%_x86_x64_%downbranch%\package.info
	echo:
	echo:
    timeout /t 7
    goto:Office16VnextInstall
:================================================================================================================
::===============================================================================================================
:wgeterror
	set "errortrigger=0"
	powershell -command "%pswindowtitle%"; Write-Host "*** ERROR downloading: %1" -foreground "Red"
	echo:
	set /p errortrigger=Cancel Download now? (1/0) ^>
	if "%errortrigger%"=="1" (rd "%downpath%\%o16build%_%o16lang%_x86_x64_%downbranch%" /S /Q)&&(goto:Office16VnextInstall)
	echo:
	goto :eof
:================================================================================================================
::===============================================================================================================
:DownloadO16Online
    pushd "%o2016rtoolpath%"
    cls
	echo:
	echo ============ DOWNLOAD OFFICE 2016 ONLINE WEB INSTALLER ===========
	echo __________________________________________________________________
    if "%o16version%"=="" set "o16version=ProPlusRetail"
    if "%o16arch%"=="" set "o16arch=x86"
    if "%o16lang%"=="" set "o16lang=en-us"
    set "of16install=0"
    set "pr16install=0"
    set "vi16install=0"
    set "installtrigger=X"
	echo:
    set /p of16install=Set Office Install (1/0) ^>
	if "%of16install%"=="1" goto:websummary
    echo:
    set /p pr16install=Set Project Install (1/0) ^>
	if "%pr16install%"=="1" goto:websummary
    echo:
    set /p vi16install=Set Visio Install (1/0) ^>
	if "%vi16install%"=="1" goto:websummary
	goto:DownloadO16Online
::===============================================================================================================
:websummary
    if "%of16install%"=="1" ((set "of16=YES")&&(set "o16version=ProPlusRetail")) else (set "of16=NO")
    if "%pr16install%"=="1" ((set "pr16=YES")&&(set "o16version=ProjectProRetail")) else (set "pr16=NO")
    if "%vi16install%"=="1" ((set "vi16=YES")&&(set "o16version=VisioProRetail")) else (set "vi16=NO")
    echo __________________________________________________________________
    echo:
	if "%winx%"=="win_x32" ((set "o16arch=x86")&&(goto:weblangselect))
    set /p o16arch=Set Architecture to install (x86 or x64) ^>
    echo:
::===============================================================================================================
:weblangselect
::===============================================================================================================
    call :seto16language
::===============================================================================================================
    echo:
    echo __________________________________________________________________
	echo:
    echo Pending Online WEB Install (SUMMARY)
    echo:
    echo Install Office ?  : %of16install% = %of16%
    echo:
    echo Install Project ? : %pr16install% = %pr16%
    echo:
    echo Install Visio ?   : %vi16install% = %vi16%
    echo:
    echo:
    echo Install Architecture ?: %o16arch%
    echo Install Language ?    : %o16lang%
    echo __________________________________________________________________
	echo:
    set /p installtrigger=Start Online WEB Install now (1/0)? ^>
    if "%installtrigger%"=="0" goto:Office16VnextInstall
    if "%installtrigger%"=="1" goto:OfficeWebInstall
    goto:DownloadO16Online
:================================================================================================================
::===============================================================================================================
:OfficeWebInstall
    if "%o16version%"=="ProPlusRetail" set "o16gvlk=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    if "%o16version%"=="ProjectProRetail" set "o16gvlk=YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
    if "%o16version%"=="VisioProRetail" set "o16gvlk=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
    cls
	echo:
    echo ============ DOWNLOAD OFFICE 2016 ONLINE INSTALLER ===============
	echo __________________________________________________________________
	echo:
    echo Sending generated link to your browser.
    echo:
    echo Save the offered Setup.exe and run it to start Online WEB Install
    start "" "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=%o16version%&language=%o16lang%&platform=%o16arch%&token=%o16gvlk%&version=O16GA&act=1"
    echo __________________________________________________________________
	echo:
    echo:
	timeout /t 7
    goto:Office16VnextInstall
:================================================================================================================
::===============================================================================================================
:InstallO16
	del /f /q "%downpath%\wget.exe" >nul 2>&1
	set "o16build="
	set "mo16install=0"
	set "of16install=0"
	set "of36install=0"
	set "st16install=0"
	set "wd16disable=0"
	set "ex16disable=0"
	set "pp16disable=0"
	set "ac16disable=0"
	set "ol16disable=0"
	set "pb16disable=0"
	set "on16disable=0"
	set "sk16disable=0"
	set "od16disable=0"
	set "sd16disable=0"
	set "ip16disable=0"
	set "pr16install=0"
	set "vi16install=0"
	set "wd16install=0"
	set "ex16install=0"
	set "pp16install=0"
	set "ac16install=0"
	set "ol16install=0"
	set "pb16install=0"
	set "on16install=0"
	set "sk16install=0"
	set "installtrigger=X"
	set "createpackage=0"
	cls
	echo:
	echo ============ INSTALL VNEXT OFFICE 2016 ===========================
	echo __________________________________________________________________
	echo:
	if "%downpath%"=="not set" set /p downpath=Set VNEXT Download Path ^>
	set "downpath=%downpath:"=%"
	set "downdrive=%downpath:~0,2%"
	cd /d %downdrive% >nul 2>&1
	if errorlevel 1 (echo:)&&(echo Unknown Drive "%downdrive%" - Drive not found)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:InstallO16)
	if "%downpath:~-1%"=="\" set "downpath=%downpath:~0,-1%"
	pushd %o2016rtoolpath%
	echo:
	echo List of available installation packages
	echo:
	echo #   Package
	set /a countx=0
	for /f "tokens=*" %%a in ('dir "%downpath%" /ad /b 2^>nul ^| findstr /i "16.0"') do (
	echo:
	SET /a countx=!countx! + 1
	set packagelist!countx!=%%a
	echo !countx!   %%a
	)
	if %countx% EQU 0 ((echo No install packages found in: "%downpath%")&&(echo.)&&(pause)&&(set "downpath=not set")&&(goto:InstallO16))
	echo:
	echo:
	set /a packnum=0
	set /p packnum=Enter package number # ^>
	if %packnum% EQU 0 goto:InstallO16
	if %packnum% GTR %countx% goto:InstallO16
	echo:
	set "installpath=%downpath%\!packagelist%packnum%!"
	if "%installpath:~-1%"=="\" set "installpath=%installpath:~0,-1%"
	set countx=0
	for /f "tokens=*" %%a in (%installpath%\package.info) do (
    SET /a countx=!countx! + 1
    set var!countx!=%%a
	)
	if %countx% EQU 0 (echo:)&&(pause)&&(set "downpath=not set")&&(goto:InstallO16)
	set "o16build=%var1%"
	set "o16lang=%var2%"
	set "o16updlocid=%var3%"
::===============================================================================================================
	cls
	echo:
	echo ============ INSTALL VNEXT OFFICE 2016 ===========================
	echo __________________________________________________________________
	echo:
	echo Using VNEXT Office Setup Package found in:
	echo %installpath%
	echo:
	set /p mo16install=Set Office Mondo Install (1/0) ^>
	if "%mo16install%"=="1" (goto:InstallExclusions)
	set /p of16install=Set Office ProfessionalPlus Install (1/0) ^>
	if "%of16install%"=="1" (goto:InstallExclusions)
	set /p of36install=Set Office 365 Install (1/0) ^>
	if "%of36install%"=="1" (goto:InstallExclusions)
	set /p st16install=Set Office Professional/Standard Install (1/0) ^>
	if "%st16install%"=="1" (goto:InstallExclusions)
::===============================================================================================================
	echo:
	set /p wd16install=Set Word Single App Install (1/0) ^>
	set /p ex16install=Set Excel Single App Install (1/0) ^>
	set /p pp16install=Set Powerpoint Single App Install (1/0) ^>
	set /p ac16install=Set Access Single App Install (1/0) ^>
	set /p ol16install=Set Outlook Single App Install (1/0) ^>
	set /p pb16install=Set Publisher Single App Install (1/0) ^>
	set /p on16install=Set OneNote Single App Install (1/0) ^>
	set /p sk16install=Set Skype For Business Single App Install (1/0) ^>
	goto:InstallProVis
::===============================================================================================================
:InstallExclusions
	if "%mo16install%"=="1" ((set "of16install=0")&&(set "of36install=0")&&(set "st16install=0"))
	if "%of16install%"=="1" ((set "mo16install=0")&&(set "of36install=0")&&(set "st16install=0"))
	if "%of36install%"=="1" ((set "mo16install=0")&&(set "of16install=0")&&(set "st16install=0"))
	if "%st16install%"=="1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of36install=0"))
	echo:
	echo Full Suite Install Exclusion List - Disable not needed Office Programs
	set /p wd16disable=Disable Word Install  (1/0) ^>
	set /p ex16disable=Disable Excel Install (1/0) ^>
	set /p pp16disable=Disable Powerpoint Install (1/0) ^>
	set /p ac16disable=Disable Access Install (1/0) ^>
	set /p ol16disable=Disable Outlook Install (1/0) ^>
	set /p pb16disable=Disable Publisher Install (1/0) ^>
	set /p on16disable=Disable OneNote Install (1/0) ^>
	set /p sk16disable=Disable Skype For Business Install (1/0) ^>
	set /p od16disable=Disable OneDrive For Business Install (1/0) ^>
	set /p sd16disable=Disable SharePointDesigner Install (1/0) ^>
	set /p ip16disable=Disable InfoPath Install (1/0) ^>
	if "%mo16install%"=="1" ((set "of16install=0")&&(set "of36install=0")&&(set "st16install=0"))
	if "%of16install%"=="1" ((set "mo16install=0")&&(set "of36install=0")&&(set "st16install=0"))
	if "%of36install%"=="1" ((set "mo16install=0")&&(set "of16install=0")&&(set "st16install=0"))
	if "%st16install%"=="1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of36install=0"))
::===============================================================================================================
:InstallProVis
	echo:
	set /p pr16install=Set Project Install (1/0) ^>
	set /p vi16install=Set Visio Install (1/0) ^>
::===============================================================================================================
	echo:
	set "o16arch=x86"
	set /p o16arch=Set Architecture (x86 or x64) ^>
	if "%winx%"=="win_x32" set "o16arch=x86"
::===============================================================================================================
	echo __________________________________________________________________
	echo:
	if "%o16updlocid%"=="64256afe-f5d9-4f86-8936-8840a6a4f5be" ((echo == Office %o16build% Insider Branch Pending Setup -SUMMARY- ==)&&(goto:PendSetupContinue))
	if "%o16updlocid%"=="492350f6-3a01-4f97-b9c0-c7c6ddf67d60" ((echo == Office %o16build% Current Branch Pending Setup -SUMMARY- ==)&&(goto:PendSetupContinue))
	echo == Office %o16build% Manual Override Branch Pending Setup -SUMMARY- ===
:PendSetupContinue
	echo:
	echo The following programs are selected for install:
	echo:
	if "%wd16install%"=="1" goto:PendSetupSingleApp
	if "%ex16install%"=="1" goto:PendSetupSingleApp
	if "%pp16install%"=="1" goto:PendSetupSingleApp
	if "%ac16install%"=="1" goto:PendSetupSingleApp
	if "%ol16install%"=="1" goto:PendSetupSingleApp
	if "%pb16install%"=="1" goto:PendSetupSingleApp
	if "%on16install%"=="1" goto:PendSetupSingleApp
	if "%sk16install%"=="1" goto:PendSetupSingleApp
::===============================================================================================================
	if "%mo16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Mondo 2016 Grande Suite" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Office 2016 Professional Plus" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of36install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Office 365 Professional Plus" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%st16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Office 2016 Professional/Standard" -foreground "Green")&&(goto:PendSetupFullSuite)
	goto:PendSetupProjectVisio
:PendSetupFullSuite
	if "%wd16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Word" -foreground "Red")
	if "%wd16disable%"=="0" (echo --^> Enabled:  Word)
	if "%ex16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Excel" -foreground "Red")
	if "%ex16disable%"=="0" (echo --^> Enabled:  Excel)
	if "%pp16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Powerpoint" -foreground "Red")
	if "%pp16disable%"=="0" (echo --^> Enabled:  PowerPoint)
	if "%ac16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Access" -foreground "Red")
	if "%ac16disable%"=="0" (echo --^> Enabled:  Access)
	if "%ol16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Outlook" -foreground "Red")
	if "%ol16disable%"=="0" (echo --^> Enabled:  Outlook)
	if "%pb16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Publisher" -foreground "Red")
	if "%pb16disable%"=="0" (echo --^> Enabled:  Publisher)
	if "%on16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: OneNote" -foreground "Red")
	if "%on16disable%"=="0" (echo --^> Enabled:  OneNote)
	if "%sk16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: Skype For Business" -foreground "Red")
	if "%sk16disable%"=="0" (echo --^> Enabled:  Skype For Business)
	if "%od16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: OneDrive For Business" -foreground "Red")
	if "%od16disable%"=="0" (echo --^> Enabled:  OneDrive For Business)
	if "%sd16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: SharePointDesigner" -foreground "Red")
	if "%sd16disable%"=="0" (echo --^> Enabled:  SharePointDesigner)
	if "%ip16disable%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "--> Disabled: InfoPath" -foreground "Red")
	if "%ip16disable%"=="0" (echo --^> Enabled:  InfoPath)
	echo:
	goto:PendSetupProjectVisio
::===============================================================================================================
:PendSetupSingleApp	
	if "%wd16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Word 2016 Single App" -foreground "Green")
	if "%ex16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Excel 2016 Single App" -foreground "Green")
	if "%pp16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "PowerPoint 2016 Single App" -foreground "Green")
	if "%ac16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Access 2016 Single App" -foreground "Green")
	if "%ol16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Outlook 2016 Single App" -foreground "Green")
	if "%pb16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Publisher 2016 Single App" -foreground "Green")
	if "%on16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "OneNote 2016 Single App" -foreground "Green")
	if "%sk16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Skype For Business 2016 Single App" -foreground "Green")
	echo:
::===============================================================================================================
:PendSetupProjectVisio
	if "%pr16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Project 2016 Professional" -foreground "Green")
	if "%vi16install%"=="1" (powershell -command "%pswindowtitle%"; Write-Host "Visio 2016 Professional" -foreground "Green")
::===============================================================================================================
	echo:
	echo Architecture: %o16arch%
::===============================================================================================================
    echo Language:     %o16lang%   (fixed - matches vnext download package)
	echo __________________________________________________________________
	echo:
	set /p installtrigger=Start local install now (1/0) or Create Install Package (C) ? ^>
	if "%installtrigger%"=="0" goto:Office16VnextInstall
	if "%installtrigger%"=="C" set "createpackage=1"
	if "%installtrigger%"=="c" set "createpackage=1"
	if "%installtrigger%"=="1" goto:OfficeInstallXML
	if "%createpackage%"=="1" goto:OfficeInstallXML
	goto:InstallO16
:================================================================================================================
::===============================================================================================================
:OfficeInstallXML
	cls
    echo:
	echo ============ INSTALL VNEXT OFFICE ================================
	echo __________________________________________________________________
    echo:
    if "%o16arch%"=="x64" (set "o16a=64") else (set "o16a=32")
	echo Creating setup files "setup.exe", "configure%o16a%.xml" and "start_setup.cmd"
	echo:
    echo in Installpath "%installpath%"
    echo:
	set "oxml=%installpath%\configure%o16a%.xml"
	set "obat=%installpath%\start_setup.cmd"
	if exist "%installpath%\configure*.xml" del /s /q "%installpath%\configure*.xml" >nul 2>&1
	if exist "%obat%" del /s /q "%obat%" >nul 2>&1
	echo D|xcopy "%o2016rtoolpath%\OfficeFixes\setup.exe" "%installpath%" /s /q /y >nul 2>&1
    echo ^<Configuration^> >"%oxml%"
    echo     ^<Add OfficeClientEdition="%o16a%"^> >>"%oxml%"
    if "%mo16install%"=="1" (
        echo         ^<Product ID="MondoRetail"^> >>"%oxml%"
        echo            ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%"=="1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%"=="1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%"=="1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%"=="1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%"=="1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%"=="1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%"=="1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%"=="1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%"=="1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
		if "%sd16disable%"=="1" echo             ^<ExcludeApp ID="SharePointDesigner"/^> >>"%oxml%"
		if "%ip16disable%"=="1" echo             ^<ExcludeApp ID="InfoPath"/^> >>"%oxml%"
        echo         ^</Product^>/^> >>"%oxml%"
	)
    if "%of16install%"=="1" (
        echo         ^<Product ID="ProplusRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%"=="1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%"=="1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%"=="1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%"=="1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%"=="1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%"=="1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%"=="1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%"=="1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%"=="1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%sd16disable%"=="1" echo             ^<ExcludeApp ID="SharePointDesigner"/^> >>"%oxml%"
		if "%ip16disable%"=="1" echo             ^<ExcludeApp ID="InfoPath"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%of36install%"=="1" (
        echo         ^<Product ID="O365ProplusRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%"=="1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%"=="1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%"=="1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%"=="1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%"=="1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%"=="1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%"=="1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%"=="1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%"=="1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%sd16disable%"=="1" echo             ^<ExcludeApp ID="SharePointDesigner"/^> >>"%oxml%"
		if "%ip16disable%"=="1" echo             ^<ExcludeApp ID="InfoPath"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%st16install%"=="1" (
        echo         ^<Product ID="ProfessionalRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
		if "%wd16disable%"=="1" echo             ^<ExcludeApp ID="Word"/^> >>"%oxml%"
		if "%ex16disable%"=="1" echo             ^<ExcludeApp ID="Excel"/^> >>"%oxml%"
		if "%pp16disable%"=="1" echo             ^<ExcludeApp ID="PowerPoint"/^> >>"%oxml%"
		if "%ac16disable%"=="1" echo             ^<ExcludeApp ID="Access"/^> >>"%oxml%"
		if "%ol16disable%"=="1" echo             ^<ExcludeApp ID="Outlook"/^> >>"%oxml%"
		if "%pb16disable%"=="1" echo             ^<ExcludeApp ID="Publisher"/^> >>"%oxml%"
		if "%on16disable%"=="1" echo             ^<ExcludeApp ID="OneNote"/^> >>"%oxml%"
		if "%sk16disable%"=="1" echo             ^<ExcludeApp ID="Lync"/^> >>"%oxml%"
		if "%od16disable%"=="1" echo             ^<ExcludeApp ID="Groove"/^> >>"%oxml%"
        if "%sd16disable%"=="1" echo             ^<ExcludeApp ID="SharePointDesigner"/^> >>"%oxml%"
		if "%ip16disable%"=="1" echo             ^<ExcludeApp ID="InfoPath"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pr16install%"=="1" (
        echo         ^<Product ID="ProjectProRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%vi16install%"=="1" (
        echo         ^<Product ID="VisioProRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%wd16install%"=="1" (
        echo         ^<Product ID="WordRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ex16install%"=="1" (
        echo         ^<Product ID="ExcelRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pp16install%"=="1" (
        echo         ^<Product ID="PowerPointRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ac16install%"=="1" (
        echo         ^<Product ID="AccessRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%ol16install%"=="1" (
        echo         ^<Product ID="OutlookRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%pb16install%"=="1" (
        echo         ^<Product ID="PublisherRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%on16install%"=="1" (
        echo         ^<Product ID="OneNoteRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    if "%sk16install%"=="1" (
        echo         ^<Product ID="SkypeForBusinessRetail"^> >>"%oxml%"
        echo             ^<Language ID="%o16lang%"/^> >>"%oxml%"
        echo         ^</Product^> >>"%oxml%"
	)
    echo     ^</Add^> >>"%oxml%"
    echo     ^<Display Level="Full" AcceptEULA="TRUE" /^> >>"%oxml%"
	echo     ^<Updates Enabled="TRUE" UpdatePath="http://officecdn.microsoft.com/pr/%o16updlocid%" /^> >>"%oxml%"
	echo ^</Configuration^> >>"%oxml%"
    echo setup.exe /configure configure%o16a%.xml >"%obat%"
    if not "%createpackage%"=="1" (echo:)&&(echo Setup is starting...)&&(start "" /MIN "%installpath%\setup.exe" /configure "%oxml%")
    echo:
    echo __________________________________________________________________
    echo:
	echo:
	timeout /t 7
    goto:Office16VnextInstall
:================================================================================================================
::===============================================================================================================
:CheckOfficeApplications
	set "_MondoRetail="
	set "_ProPlusRetail="
	set "_ProfessionalRetail="
	set "_ProjectProRetail="
	set "_VisioProRetail="
	set "_WordRetail="
	set "_ExcelRetail="
	set "_PowerPointRetail="
	set "_AccessRetail="
	set "_OutlookRetail="
	set "_PublisherRetail="
	set "_OneNoteRetail="
	set "_SkypeForBusinessRetail="
	set "ProPlusVLFound="
	set "StandardVLFound="
	set "ProjectProVLFound="
	set "VisioProVLFound="
	set "installpath16="
	set "officepath3="
	set "o16arch="
	reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOffice16C2R
	reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOfficeVL32onW64
	reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" >nul 2>&1
	if %errorlevel% EQU 0 goto:CheckOfficeVL32W32orVL64W64
	goto:Office16VnextInstall
::===============================================================================================================
:CheckOffice16C2R
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "Platform" 2^>nul') DO (set "o16arch=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "VersionToReport" 2^>nul') DO (set "o16version=%%B")
	for /f "tokens=3,4 delims=." %%A in ("%o16version%") do (set /a offvermajor=%%A && set /a offverminor=%%B)
	set "o16updlocid=64256afe-f5d9-4f86-8936-8840a6a4f5be"
	if %offvermajor% EQU 4229 if %offverminor% EQU 1015 set "o16updlocid=a1d255fe-49bb-4a88-ade8-4ec08d98fbd4"
	if %offvermajor% EQU 4266 if %offverminor% EQU 1003 set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
	reg add "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /f /v "CDNBaseUrl" /t REG_SZ /d %o16updlocid% >nul 2>&1
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" 2^>nul') DO (Set "installpath16=%%B")
	set "officepath3=%installpath16%\Office16"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "ProductReleaseIds" 2^>nul') DO (Set "Office16AppsInstalled=%%B")
	for /F "tokens=1,2,3,4,5,6,7,8,9,10,11,12,13 delims=," %%A IN ("%Office16AppsInstalled%") DO (
	set "_%%A=YES"
	set "_%%B=YES"
	set "_%%C=YES"
	set "_%%D=YES"
	set "_%%E=YES"
	set "_%%F=YES"
	set "_%%G=YES"
	set "_%%H=YES"
	set "_%%I=YES"
	set "_%%J=YES"
	set "_%%K=YES"
	set "_%%L=YES"
	set "_%%M=YES"
	)
	goto:eof
::===============================================================================================================
:CheckOfficeVL32onW64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound%" EQU "Microsoft Office Professional Plus 2016" set "_ProPlusRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound%" EQU "Microsoft Office Standard 2016" set "_ProfessionalRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound%" EQU "Microsoft Project Professional 2016" set "_ProjectProRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound%" EQU "Microsoft Visio Professional 2016" set "_VisioProRetail=YES"
	if "%_ProPlusRetail%"=="YES" goto:OfficeVL32onW64Path
	if "%_ProfessionalRetail%"=="YES" goto:OfficeVL32onW64Path
	if "%_ProjectProRetail%"=="YES" goto:OfficeVL32onW64Path
	if "%_VisioProRetail%"=="YES" goto:OfficeVL32onW64Path
	goto:Office16VnextInstall
:OfficeVL32onW64Path
	set "o16arch=x86"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") >nul 2>&1
	set "officepath3=%installpath16%"
	goto:eof
:CheckOfficeVL32W32orVL64W64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound%" EQU "Microsoft Office Professional Plus 2016" set "_ProPlusRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound%" EQU "Microsoft Office Standard 2016" set "_ProfessionalRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound%" EQU "Microsoft Project Professional 2016" set "_ProjectProRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound%" EQU "Microsoft Visio Professional 2016" set "_VisioProRetail=YES"
	if "%_ProPlusRetail%"=="YES" (set "o16arch=x86")&&(goto:OfficeVL32V64Path)
	if "%_ProfessionalRetail%"=="YES" (set "o16arch=x86")&&(goto:OfficeVL32V64Path)
	if "%_ProjectProRetail%"=="YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	if "%_VisioProRetail%"=="YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") >nul 2>&1
	if "%ProPlusVLFound%" EQU "Microsoft Office Professional Plus 2016" set "_ProPlusRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") >nul 2>&1
	if "%StandardVLFound%" EQU "Microsoft Office Standard 2016" set "_ProfessionalRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") >nul 2>&1
	if "%ProjectProVLFound%" EQU "Microsoft Project Professional 2016" set "_ProjectProRetail=YES"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") >nul 2>&1
	if "%VisioProVLFound%" EQU "Microsoft Visio Professional 2016" set "_VisioProRetail=YES"
	if "%_ProPlusRetail%"=="YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_ProfessionalRetail%"=="YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_ProjectProRetail%"=="YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	if "%_VisioProRetail%"=="YES" (set "o16arch=x64")&&(goto:OfficeVL32V64Path)
	goto:Office16VnextInstall
:OfficeVL32VL64Path
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") >nul 2>&1
	set "officepath3=%installpath16%"
	goto:eof
:================================================================================================================
::===============================================================================================================
:Convert16Activate
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	echo ============ CONVERT VNEXT AND WEB OFFICE ========================
    echo __________________________________________________________________
    echo:
	echo:
	echo ============ Installed Office Applications (SUMMARY) =============
    echo:
	echo:
	echo Installation path  = "%installpath16%"
	echo:
	echo:
	if defined _MondoRetail (echo Mondo Grande Suite = "%_MondoRetail%") else (echo Mondo Grande Suite = "NO")
	echo:
	if defined _ProPlusRetail (echo Professional-Plus  = "%_ProPlusRetail%") else (echo Professional-Plus  = "NO")
	echo:
	if defined _ProfessionalRetail (echo Profess./Standard  = "%_ProfessionalRetail%") else (echo Profess./Standard  = "NO")
	echo:
	if defined _VisioProRetail (echo VisioPro           = "%_VisioProRetail%") else (echo VisioPro           = "NO")
	echo:
	if defined _ProjectProRetail (echo ProjectPro         = "%_ProjectProRetail%") else (echo ProjectPro         = "NO")
	echo:
	if defined _WordRetail (echo Word                = "%_WordRetail%") else (echo Word               = "NO")
	echo:
	if defined _ExcelRetail (echo Excel               = "%_ExcelRetail%") else (echo Excel              = "NO")
	echo:
	if defined _PowerPointRetail (echo PowerPoint          = "%_PowerPointRetail%") else (echo PowerPoint         = "NO")
	echo:
	if defined _AccessRetail (echo Access              = "%_AccessRetail%") else (echo Access             = "NO")
	echo:
	if defined _OutlookRetail (echo Outlook             = "%_OutlookRetail%") else (echo Outlook            = "NO")
	echo:
	if defined _PublisherRetail (echo Publisher           = "%_PublisherRetail%") else (echo Publisher          = "NO")
	echo:
	if defined _OneNoteRetail (echo OneNote             = "%_OneNoteRetail%") else (echo OneNote            = "NO")
	echo:
	if defined _SkypeForBusinessRetail (echo Skype               = "%_SkypeForBusinessRetail%") else (echo Skype              = "NO")
	echo:
	echo __________________________________________________________________
	echo:
	echo:
	pause
::===============================================================================================================
	set "VolumeMode=0"
	set "CleanupMode=0"
	cls
	echo:
	echo ============== CHECKING OFFICE 2016 LICENSES ================================
	echo _____________________________________________________________________________
	echo:
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "VolumeMode=1"
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%TIMEBASED%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if %win% GEQ 9200 wmic path %slp% where (Description like '%%Grace%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if "%CleanupMode%"=="0" if "%VolumeMode%"=="1" goto:skip_cleanup
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%KMSCLIENT%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "VolumeMode=1"
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%TIMEBASED%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if %win% LSS 9200 wmic path %ospp% where (Description like '%%Grace%%') get Name /format:list 2>nul | findstr /i /C:"Office 16" 1>nul && set "CleanupMode=1"
	if "%CleanupMode%"=="0" if "%VolumeMode%"=="1" goto:skip_cleanup
::===============================================================================================================
	echo:
	echo Cleanup - Removing Current Office 2016 - Trial-/Grace-Licenses
	echo:
	"%o2016rtoolpath%\OfficeFixes\%winx%\cleanospp.exe" -PKey
	echo:
::===============================================================================================================
	echo __________________________________________________________________
	echo:
	timeout /t 7
	cls
	echo ============ CONVERT VNEXT AND WEB OFFICE ========================
	echo __________________________________________________________________
	echo:
	if not exist "%officepath3%\ospp.vbs" xcopy "OfficeFixes\ospp\*.*" "%officepath3%" /C /E /I /Q /R /S /Y  >nul 2>&1
	call :Office16ConversionLoop
	goto:Office16VnextInstall
::===============================================================================================================
:skip_cleanup
	echo:
	echo All OK. No Conversion or Cleanup required.
	echo:
	echo __________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:Office16VnextInstall
:================================================================================================================
::===============================================================================================================
:StartKMSActivation
::===============================================================================================================
	set "kmstrigger=X"
	set "O2016RToolKMS=YES"
	set "ExtLocKMS=X"
	set /a num1=%random% %% 50+30
	set /a num2=%random% %% 186+20
	set "Host=10.%num1%.3.%num2%"
	if %win% LEQ 9200 set Host=127.0.0.2
	set "Port=16880"
	set "AI=120"
	set "RI=10080"
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	echo =========== KMS ACTIVATION VNEXT AND WEB OFFICE ==================
	echo __________________________________________________________________
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Activation" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
	if %win% GEQ 9600 reg add "HKLM\SOFTWARE\Classes\AppID\slui.exe" /f /v "NoGenTicket" /d 1 /t "REG_DWORD" >nul 2>&1
::===============================================================================================================
	echo:
	echo: 
	goto:LocalKMS
	set /p kmstrigger=Use internal O2016RTool KMS server (1/0)? ^>
    if "%kmstrigger%"=="0" goto:ChangeKMSParameter
    if "%kmstrigger%"=="1" goto:LocalKMS
    goto:StartKMSActivation
::===============================================================================================================
:ChangeKMSParameter
	set "O2016RToolKMS=NO"
	echo:
	set /p ExtLocKMS=Use (E)xternal KMS on WAN/LAN or other (L)ocal KMS server (E/L)? ^>
	echo:
	set /p Host=Set KMS host IP ^>
	echo:
	set /p Port=Set KMS host PORT ^>
	if "%ExtLocKMS%"=="E" goto:KMSExtern1
	if "%ExtLocKMS%"=="e" goto:KMSExtern1
	if "%ExtLocKMS%"=="L" goto:LocalKMS
	if "%ExtLocKMS%"=="l" goto:LocalKMS
	goto:StartKMSActivation
::===============================================================================================================
:LocalKMS
	call :StopService "sppsvc"
	if defined ospsversion (call :StopService "osppsvc")
	xcopy "%o2016rtoolpath%\OfficeFixes\%winx%\SppExtComObjPatcher.exe" "%SystemRoot%\system32\" /Y /Q >nul 2>&1
	xcopy "%o2016rtoolpath%\OfficeFixes\%winx%\SppExtComObjHook.dll" "%SystemRoot%\system32\" /Y /Q >nul 2>&1
	if "%O2016RToolKMS%"=="YES" xcopy "%o2016rtoolpath%\OfficeFixes\%winx%\vlmcsd.exe" "%SystemRoot%\system32\" /Y /Q >nul 2>&1
	call :CreateIFEOEntry "SppExtComObj.exe"
	if defined ospsversion (call :CreateIFEOEntry "osppsvc.exe")
::===============================================================================================================
:KMSExtern1
	echo:
	echo __________________________________________________________________
	echo:
	if defined _MondoRetail call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
	if defined _ProPlusRetail call :Office16Activate d450596f-894d-49e0-966a-fd39ed4c4c64
	if defined _ProfessionalRetail call :Office16Activate dedfa23d-6ed1-45a6-85dc-63cae0546de6
	if defined _VisioProRetail call :Office16Activate 6bf301c1-b94a-43e9-ba31-d494598c47fb
	if defined _ProjectProRetail call :Office16Activate 4f414197-0fc2-4c01-b68a-86cbb9ac254c
	if defined _WordRetail call :Office16Activate bb11badf-d8aa-470e-9311-20eaf80fe5cc
	if defined _ExcelRetail call :Office16Activate c3e65d36-141f-4d2f-a303-a842ee756a29
	if defined _PowerPointRetail call :Office16Activate d70b1bba-b893-4544-96e2-b7a318091c33
	if defined _AccessRetail call :Office16Activate 67c0fc0c-deba-401b-bf8b-9c8ad8395804
	if defined _OutlookRetail call :Office16Activate ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
	if defined _PublisherRetail call :Office16Activate 041a06cb-c5b8-4772-809f-416d03d16654
	if defined _OneNoteRetail call :Office16Activate d8cace59-33d2-4ac7-9b1b-9b72339c51c8
	if defined _SkypeForBusinessRetail call :Office16Activate 83e04ee1-fa8d-436d-8994-d31a862cab77
::===============================================================================================================
	if "%ExtLocKMS%"=="E" goto:KMSExtern2
	if "%ExtLocKMS%"=="e" goto:KMSExtern2
	call :StopService "sppsvc"
	if defined ospsversion (call :StopService "osppsvc")
	del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe" >nul 2>&1
	del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll" >nul 2>&1
	if "%O2016RToolKMS%"=="YES" del /f /q "%SystemRoot%\system32\vlmcsd.exe" >nul 2>&1
	call :RemoveIFEOEntry "SppExtComObj.exe"
	if defined ospsversion (call :RemoveIFEOEntry "osppsvc.exe")
::===============================================================================================================
:KMSExtern2
	if %win% GEQ 9600 reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% GEQ 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% GEQ 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% LSS 9600 reg delete "HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
	if %win% LSS 9600 reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
::===============================================================================================================
	sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
	echo:
	echo:
	timeout /t 7
    goto:Office16VnextInstall
:================================================================================================================
::===============================================================================================================
:Office16Activate
	set /a "GraceMin=0"
	if %win% GEQ 9200	(
	wmic path %slp% where ID="%1" call SetKeyManagementServiceMachine %Host% >nul 2>&1
	wmic path %slp% where ID="%1" call SetKeyManagementServicePort %Port% >nul 2>&1
	for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where ID="%1" get Name /value"') do (echo: && echo Activating %%a)
	wmic path %slp% where ID="%1" call Activate >nul 2>&1
	for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where ID="%1" get GracePeriodRemaining /value"') do (set /a "GraceMin=%%a")
	wmic path %slp% where ID="%1" call ClearKeyManagementServiceMachine >nul 2>&1
	wmic path %slp% where ID="%1" call ClearKeyManagementServicePort >nul 2>&1
	)
	if %win% LSS 9200	(
	wmic path %ospp% where ID="%1" call SetKeyManagementServiceMachine %Host% >nul 2>&1
	wmic path %ospp% where ID="%1" call SetKeyManagementServicePort %Port% >nul 2>&1
	for /f "tokens=2 delims==" %%a in ('"wmic path %ospp% where ID="%1" get Name /value"') do (echo: && echo Activating %%a)
	wmic path %ospp% where ID="%1" call Activate >nul 2>&1
	for /f "tokens=2 delims==" %%a in ('"wmic path %ospp% where ID="%1" get GracePeriodRemaining /value"') do (set /a "GraceMin=%%a")
	wmic path %ospp% where ID="%1" call ClearKeyManagementServiceMachine >nul 2>&1
	wmic path %ospp% where ID="%1" call ClearKeyManagementServicePort >nul 2>&1
	)
	if %GraceMin% EQU 259200 (echo Activation successful && GOTO :TheEndIsNear) else (echo Activation failed)
	echo __________________________________________________________________
	goto:eof
:================================================================================================================
::===============================================================================================================
:seto16language
    echo Possible Language VALUES:
    echo ar-SA, cs-CZ, da-DK, de-DE, el-GR, en-US, es-ES, et-EE, fi-FI, fr-FR, he-IL,
    echo hi-IN, hu-HU, it-IT, ja-JP, ko-KR, lt-LT, lv-LV, nb-NO, nl-NL, pl-PL, pt-BR,
    echo pt-PT, ro-RO, ru-RU, sv-SE, th-TH, tr-TR, uk-UA, vi-VN, zh-CN, zh-TW
    echo:
    set "o16lcid="
    set /p o16lang=Set Language ^>
    if "%o16lang%"=="ar-SA" set "o16lcid=1025"
    if "%o16lang%"=="zh-CN" set "o16lcid=2052"
    if "%o16lang%"=="zh-TW" set "o16lcid=1028"
    if "%o16lang%"=="cs-CZ" set "o16lcid=1029"
    if "%o16lang%"=="da-DK" set "o16lcid=1030"
    if "%o16lang%"=="nl-NL" set "o16lcid=1043"
    if "%o16lang%"=="en-US" set "o16lcid=1033"
    if "%o16lang%"=="et-EE" set "o16lcid=1061"
    if "%o16lang%"=="fi-FI" set "o16lcid=1035"
    if "%o16lang%"=="fr-FR" set "o16lcid=1036"
    if "%o16lang%"=="de-DE" set "o16lcid=1031"
    if "%o16lang%"=="el-GR" set "o16lcid=1032"
    if "%o16lang%"=="he-IL" set "o16lcid=1037"
    if "%o16lang%"=="hi-IN" set "o16lcid=1081"
    if "%o16lang%"=="hu-HU" set "o16lcid=1038"
    if "%o16lang%"=="it-IT" set "o16lcid=1040"
    if "%o16lang%"=="ja-JP" set "o16lcid=1041"
    if "%o16lang%"=="ko-KR" set "o16lcid=1042"
    if "%o16lang%"=="lv-LV" set "o16lcid=1062"
    if "%o16lang%"=="lt-LT" set "o16lcid=1063"
    if "%o16lang%"=="nb-NO" set "o16lcid=1044"
    if "%o16lang%"=="pl-PL" set "o16lcid=1045"
    if "%o16lang%"=="pt-BR" set "o16lcid=1046"
    if "%o16lang%"=="pt-PT" set "o16lcid=2070"
    if "%o16lang%"=="ro-RO" set "o16lcid=1048"
    if "%o16lang%"=="ru-RU" set "o16lcid=1049"
    if "%o16lang%"=="es-ES" set "o16lcid=3082"
    if "%o16lang%"=="sv-SE" set "o16lcid=1053"
    if "%o16lang%"=="th-TH" set "o16lcid=1054"
    if "%o16lang%"=="tr-TR" set "o16lcid=1055"
    if "%o16lang%"=="uk-UA" set "o16lcid=1058"
    if "%o16lang%"=="vi-VN" set "o16lcid=1066"
    if "%o16lcid%"=="" (set "o16lang=en-US")&&(set "o16lcid=1033")
    goto:eof	
:================================================================================================================
::===============================================================================================================
:ConvertOffice16
	cls
	echo:
	echo ============ %2 2016 found ======================
	echo _____________________________________________________________________________
	echo:
	if %win% GEQ 9200    (    
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ppd.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ul.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-pl.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ppd.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ul-phn.xrm-ms"
	)
	if %win% LSS 9200    (    
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ppd.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ul.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client-ul-oob.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-pl.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ppd.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ul-oob.xrm-ms"
    cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK-ul-phn.xrm-ms"
    )
    echo __________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:eof
:================================================================================================================
::===============================================================================================================
:ConvertGeneral16
	cls
	echo ============ Office 2016 General Client found ====================
	echo __________________________________________________________________
	echo:
	if %win% GEQ 9200    (    
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office-client15.xrm-ms"
	)
	if %win% LSS 9200    (    
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%officepath3%\ospp.vbs" /inslic:"%o2016rtoolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office-client15.xrm-ms"
	)
	echo __________________________________________________________________
	echo:
	echo:
	timeout /t 7
	goto:eof
:================================================================================================================
::===============================================================================================================
:Office16ConversionLoop
	call :ConvertGeneral16
	if defined _MondoRetail call :ConvertOffice16 Mondo
	if defined _ProPlusRetail call :ConvertOffice16 ProPlus
	if defined _ProfessionalRetail call :ConvertOffice16 Standard
	if defined _ProjectProRetail call :ConvertOffice16 ProjectPro
	if defined _VisioProRetail call :ConvertOffice16 VisioPro
	if defined _WordRetail call :ConvertOffice16 Word
	if defined _ExcelRetail call :ConvertOffice16 Excel
	if defined _PowerPointRetail call :ConvertOffice16 PowerPoint
	if defined _AccessRetail call :ConvertOffice16 Access
	if defined _OutlookRetail call :ConvertOffice16 Outlook
	if defined _PublisherRetail call :ConvertOffice16 Publisher
	if defined _OneNoteRetail call :ConvertOffice16 OneNote
	if defined _SkypeForBusinessRetail call :ConvertOffice16 SkypeForBusiness
	cls
	echo ============ INSTALLING GVLK =====================================
	echo __________________________________________________________________
	if %win% GEQ 9200 if defined _MondoRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Mondo 2016 Grande Suite"
	if %win% LSS 9200 if defined _MondoRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Mondo 2016 Grande Suite"
	if %win% GEQ 9200 if defined _ProPlusRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 Professional-Plus"
	if %win% LSS 9200 if defined _ProPlusRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office 2016 Professional-Plus"
	if %win% GEQ 9200 if defined _ProfessionalRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office 2016 Professional-Standard"
	if %win% LSS 9200 if defined _ProfessionalRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office 2016 Professional-Standard"
	if %win% GEQ 9200 if defined _ProjectProRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project 2016 Professional"
	if %win% LSS 9200 if defined _ProjectProRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project 2016 Professional"
	if %win% GEQ 9200 if defined _VisioProRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio 2016 Professional"
	if %win% LSS 9200 if defined _VisioProRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio 2016 Professional"
	if %win% GEQ 9200 if defined _WordRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% LSS 9200 if defined _WordRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% GEQ 9200 if defined _ExcelRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% LSS 9200 if defined _ExcelRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% GEQ 9200 if defined _PowerPointRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% LSS 9200 if defined _PowerPointRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% GEQ 9200 if defined _AccessRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% LSS 9200 if defined _AccessRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% GEQ 9200 if defined _OutlookRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% LSS 9200 if defined _OutlookRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% GEQ 9200 if defined _PublisherRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% LSS 9200 if defined _PublisherRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% GEQ 9200 if defined _OneNoteRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% LSS 9200 if defined _OneNoteRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% GEQ 9200 if defined _SkypeForBusinessRetail call :OfficeGVLKInstall "%sls%",%slsversion%,"869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	if %win% LSS 9200 if defined _SkypeForBusinessRetail call :OfficeGVLKInstall "%osps%",%ospsversion%,"869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	echo:
	echo:
	timeout /t 7
	goto:eof
:================================================================================================================
::===============================================================================================================
:OfficeGVLKInstall
	echo:
	echo %4
    if %win% GEQ 9200    (
    wmic path %~1 where version='%~2' call InstallProductKey ProductKey="%~3" >nul 2>&1
    if %errorlevel% EQU 0 echo: & echo Successfully installed %~3 & echo:
    if %errorlevel% NEQ 0 echo: & echo Installing %~3 Failed & echo:
    )
    if %win% LSS 9200    (
    cscript "%officepath3%\ospp.vbs" /inpkey:%~3 >nul 2>&1
    if %errorlevel% EQU 0 echo: & echo Successfully installed %~3 & echo:
    if %errorlevel% NEQ 0 echo: & echo Installing %~3 Failed & echo:
    )
	echo __________________________________________________________________
	goto:eof
:================================================================================================================
::===============================================================================================================
:CreateIFEOEntry
	reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\%~1" /f /v "Debugger" /t REG_SZ /d "SppExtComObjPatcher.exe" >nul 2>&1
	reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\%~1" /f /v "KMS_Emulation" /t REG_DWORD /d 0 >nul 2>&1
::===============================================================================================================
	if "%O2016RToolKMS%"=="NO" goto:eof
	tasklist | find /i "vlmcsd.exe" >nul 2>&1
	if %errorlevel% EQU 0 goto:eof
::===============================================================================================================
	netsh advfirewall firewall add rule name="VLMCSD" dir=in program="%windir%\system32\vlmcsd.exe" localport=%Port% protocol=TCP action=allow remoteip=any >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out program="%windir%\system32\vlmcsd.exe" localport=%Port% protocol=TCP action=allow remoteip=any >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out protocol=any remoteip=65.55.58.195 action=block >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out protocol=any remoteip=65.52.98.231 action=block >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out protocol=any remoteip=65.52.98.232 action=block >nul 2>&1
	netsh advfirewall firewall add rule name="VLMCSD" dir=out protocol=any remoteip=65.52.98.233 action=block >nul 2>&1
	start "VLMCSD KMS" /D "%windir%\system32\" /MIN vlmcsd.exe -P %Port% -A %AI% -R %RI% -C %o16lcid% -De -r1 >nul 2>&1
	goto:eof
:================================================================================================================
::===============================================================================================================
:RemoveIFEOEntry
	if '%~1' NEQ 'osppsvc.exe' (reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\%~1" /f >nul 2>&1)
	if '%~1' EQU 'osppsvc.exe' (
	reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\osppsvc.exe" /f /v "Debugger" >nul 2>&1
	reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\osppsvc.exe" /f /v "KMS_Emulation" >nul 2>&1
	)
::===============================================================================================================
	if "%O2016RToolKMS%"=="NO" goto:eof
	tasklist | find /i "vlmcsd.exe" >nul 2>&1
	if %errorlevel% NEQ 0 goto:eof
	taskkill /t /f /IM "vlmcsd.exe" >nul 2>&1
::===============================================================================================================
	netsh advfirewall firewall delete rule name="VLMCSD" >nul 2>&1
	goto:eof
:================================================================================================================
::===============================================================================================================
:StopService
	sc query "%1" | findstr /i "STOPPED" >nul 2>&1
	if %errorlevel% NEQ 0 net stop "%1" /y >nul 2>&1
	sc query "%1" | findstr /i "STOPPED" >nul 2>&1
	if %errorlevel% NEQ 0 sc stop "%1" >nul 2>&1
	goto:eof
:================================================================================================================
::===============================================================================================================
:
	endlocal
    echo:
	echo:
	timeout /T 7
	exit
:================================================================================================================
::===============================================================================================================
