#RequireAdmin

# Remove Files

## Active Mode Install
Remove-Item -Path "C:\Program Files\MSEdgeRedirect\*" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files\MSEdgeRedirect" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files\Microsoft\Edge\Application\msedge_ifeo.exe" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files\Microsoft\Edge\Application\msedge_no_ifeo.exe" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge_ifeo.exe" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge_no_ifeo.exe" -Force -ErrorAction SilentlyContinue

## Service Mode Install
Remove-Item -Path "$env:LOCALAPPDATA\MSEdgeRedirect\*" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:LOCALAPPDATA\MSEdgeRedirect" -Force -ErrorAction SilentlyContinue

## Start Menu Items
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\*" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect" -Force -ErrorAction SilentlyContinue

# Remove Registry

## Active Mode Install
Remove-Item "HKLM:\Software\Classes\MSEdgeRedirect" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKLM:\Software\Classes\Applications\MSEdgeRedirect.exe" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe" -Recurse -Force -ErrorAction SilentlyContinue

## Service Mode Install
Remove-Item "HKCU:\Software\Classes\MSEdgeRedirect" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\Applications\MSEdgeRedirect.exe" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe" -Recurse -Force -ErrorAction SilentlyContinue

## Europe Mode Install
Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe" -Recurse -Force -ErrorAction SilentlyContinue

# Restore Registry

## Europe Mode Install (>= 0.8.1.0)

$HKLMRegion = Get-ItemProperty -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\DeviceRegion" -Name "OldRegion" -ErrorAction SilentlyContinue
$HKUName = Get-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "OldName" -ErrorAction SilentlyContinue
$HKUNation = Get-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "OldNation" -ErrorAction SilentlyContinue
$HKCUName = Get-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "OldName" -ErrorAction SilentlyContinue
$HKCUNation = Get-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "OldNation" -ErrorAction SilentlyContinue

If ($HKLMRegion) {
    Set-ItemProperty -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\DeviceRegion" -Name "Region" -Value $HKLMRegion
}
If ($HKUName) {
    Set-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "Name" -Value $HKUName
}
If ($HKUNation) {
    Set-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "Nation" -Value $HKUNation
}
If ($HKCUName) {
    Set-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "Name" -Value $HKCUName
}
If ($HKCUNation) {
    Set-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "Nation" -Value $HKCUNation
}

## Europe Mode Install (< 0.8.1.0)

$startString = "Read Pre-European Install Values of:"
$endString = " & "

try {
    $lines = Get-Content "$env:LOCALAPPDATA\MSEdgeRedirect\logs\Install.log"
    
    foreach ($line in $lines) {
        if ($line -match [regex]::Escape($startString)) {
            $startIndex = $line.IndexOf($startString)
            
            if ($startIndex -ge 0) {
                $extractStart = $startIndex + $startString.Length
                
                $endIndex = $line.IndexOf($endString, $extractStart)
                
                if ($endIndex -ge 0) {
                    $extractedText = $line.Substring($extractStart, $endIndex - $extractStart)
                    $extractedArray = $extractedText -split '\|'
                    if ($extractedArray.Count -ne 5) {
                        Write-Output "[Critical] Found Pre-Europe Mode Values but count is unexpected, bailing."
                        Exit
                    } else {
                        Set-ItemProperty -Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\DeviceRegion" -Name "Region" -Value $extractedArray[0]
                        Set-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "Name" -Value $extractedArray[1]
                        Set-ItemProperty -Path "HKEY_USERS\.DEFAULT\Control Panel\International\Geo" -Name "Nation" -Value $extractedArray[2]
                        Set-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "Name" -Value $extractedArray[3]
                        Set-ItemProperty -Path "HKEY_CURRENT_USER\Control Panel\International\Geo" -Name "Nation" -Value $extractedArray[4]
                    }
                } else {
                    Exit
                }
            }
        }
    }
}
catch {
    Exit
}

## Restore UCPD

Set-ItemProperty -Path "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\UCPD" -Name "Start" -Value 1
Set-ItemProperty -Path "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\UCPD" -Name "FeatureV2" -Value 2

# End of Script, Advise Reboot
Write-Output "It is Recommended that you Reboot your Computer."
Exit