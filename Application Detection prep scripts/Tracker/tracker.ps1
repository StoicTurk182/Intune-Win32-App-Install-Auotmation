$hives = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$hives | ForEach-Object {
    Get-ItemProperty $_ -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*Evonex*" } |
    Select-Object DisplayName, DisplayVersion, UninstallString, PSPath |
    Format-List
}
```

`Format-List` will stop the truncation.

**What to expect:** Since it's under `AppData\Local`, the registry entry will likely be in `HKCU` not `HKLM`, and the uninstall string will probably be a Squirrel-style command like:
```
"C:\Users\Administrator\AppData\Local\Programs\Evonex Connect\Uninstall Evonex Connect.exe" --uninstall