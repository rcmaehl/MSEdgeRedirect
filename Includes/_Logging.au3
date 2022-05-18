#include-once

#include <Date.au3>
#include <FileConstants.au3>

Global $hLogs[5] = _
	[FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppFailures.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppGeneral.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppSecurity.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\PEBIAT.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\URIFailures.log", $FO_APPEND)]

Global Enum $AppFailures, $AppGeneral, $AppSecurity, $PEBIAT, $URIFailures

Func _Log($hLog, $sMsg)

	FileWrite($hLog, _NowCalc() & " - " & $sMsg)

EndFunc