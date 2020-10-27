
Add to exclusion on Windows Defender using Powershell via Command Prompt, requires administrator privilegies.
```
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Add-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\O2016RTool (1).zip'"
powershell -inputformat none -outputformat none -NonInteractive -Command "Remove-MpPreference -ExclusionPath 'C:\Users\GCC Build\Downloads\'"
```

The Exclusion paths and files can be found here in the registry: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths
