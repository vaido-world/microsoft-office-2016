@ECHO OFF
SETLOCAL EnableDelayedExpansion
 
 
 
SET "o16version=ProPlusRetail"
SET "o16lang=en-US"
SET "o16arch=x64"
SET "o16gvlk=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
START "" "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=%o16version%&language=%o16lang%&platform=%o16arch%&token=%o16gvlk%&version=O16GA&act=1"


FOR /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" 2^>NUL') DO ( SET "installpath16=%%B" )
echo %installpath16%

FOR /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "ProductReleaseIds" 2^>nul') DO ( SET "Office16AppsInstalled=%%B" )
FOR /F "tokens=1,2,3,4,5,6,7,8,9,10,11,12,13 delims=," %%A IN ("%Office16AppsInstalled%") DO (
SET "_%%A=YES"
SET "_%%B=YES"
SET "_%%C=YES"
SET "_%%D=YES"
SET "_%%E=YES"
SET "_%%F=YES"
SET "_%%G=YES"
SET "_%%H=YES"
SET "_%%I=YES"
SET "_%%J=YES"
SET "_%%K=YES"
SET "_%%L=YES"
SET "_%%M=YES"
)

echo %Office16AppsInstalled%

pause