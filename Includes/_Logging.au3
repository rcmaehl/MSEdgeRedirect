#include-once

#include <Date.au3>
#include <FileConstants.au3>

Global $hLogs[6] = _
	[FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppFailures.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppGeneral.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppSecurity.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\Install.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\PEBIAT.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\URIFailures.log", $FO_APPEND)]

Global Enum $AppFailures, $AppGeneral, $AppSecurity, $Install, $PEBIAT, $URIFailures

Func _Log($hLog, $sMsg)

	FileWrite($hLog, _NowCalc() & " - " & $sMsg)

EndFunc

Func _LogClose($sLog = "All")
	If $sLog <> "All" Then
		FileClose($hLogs[Eval($sLog)])
	Else
		For $iLoop = 0 To UBound($hLogs) - 1
			FileClose($hLogs[$iLoop])
		Next
	EndIf
EndFunc