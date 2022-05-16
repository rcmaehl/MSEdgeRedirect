#include-once

#include "Includes\_Logging.au3"
#include "Includes\_Theming.au3"
#include "Includes\_Settings.au3"
#include "Includes\_Translation.au3"

; gwmi Win32_Process | where { $_.name -like "msedge*.exe"} | Select-Object CommandLine | Format-Table -Wrap -AutoSize | Out-File $env:LOCALAPPDATA\MSEdgeRedirect\logs\edge.txt
