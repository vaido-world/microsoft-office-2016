
Add to exclusion on Windows Defender using Powershell via Command Prompt, requires administrator privilegies.
```
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\O2016RTool (1).zip'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\'"
```

The Exclusion paths and files can be found here in the registry: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths


Task scheduler can be used to run automated activation program.
https://superuser.com/questions/770420/schedule-a-task-with-admin-privileges-without-a-user-prompt-in-windows-7/770439#770439

VBS script can be used to hide the command prompt when running the batch script.


```
"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --incognito "https://github.com/BoQsc/microsoft-office-2016/raw/BoQsc-patch-1/O2016RTool.zip"
```
