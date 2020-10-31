@ECHO OFF

SET "reg_query_output=reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "ProductReleaseIds""
FOR /F "skip=2 tokens=3" %%A IN ('%reg_query_output%') DO ( SET "Office16AppsInstalled=%%A" )
FOR %%A IN (%Office16AppsInstalled%) DO ( SET "%%A=yes" )
IF defined ProPlusRetail (
	echo %ProPlusRetail%
)




PAUSE
EXIT