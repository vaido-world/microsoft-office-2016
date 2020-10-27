@ECHO OFF

netsh advfirewall firewall add rule name="My Applicationhhook" dir=in action=allow program="C:\Windows\System32\SppExtComObjHook.dll" enable=yes

pause