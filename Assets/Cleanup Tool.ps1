#RequireAdmin

# Remove Files

## Active Mode Install
Remove-Item -Path "C:\Program Files\MSEdgeRedirect\*" -Force
Remove-Item -Path "C:\Program Files\MSEdgeRedirect" -Force
Remove-Item -Path "C:\Program Files\Microsoft\Edge\Application\msedge_no_ifeo.exe" -Force
Remove-Item -Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge_no_ifeo.exe" -Force

## Service Mode Install
Remove-Item -Path $env:LOCALAPPDATA\MSEdgeRedirect\* -Force
Remove-Item -Path $env:LOCALAPPDATA\MSEdgeRedirect -Force

## Start Menu Items
Remove-Item -Path $env:APPDATA\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\* -Force
Remove-Item -Path $env:APPDATA\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect -Force

# Remove Registry

## Active Mode Install
Remove-Item "HKLM:\Software\Classes\MSEdgeRedirect" -Recurse -Force
Remove-Item "HKLM:\Software\Classes\Applications\MSEdgeRedirect.exe" -Recurse -Force
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect" -Recurse -Force
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe" -Recurse -Force
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe" -Recurse -Force

## Service Mode Install
Remove-Item "HKCU:\Software\Classes\MSEdgeRedirect" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\Applications\MSEdgeRedirect.exe" -Recurse -Force
Remove-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect" -Recurse -Force
Remove-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe" -Recurse -Force
