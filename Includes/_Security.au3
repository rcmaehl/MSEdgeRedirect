#include-once

#include <WinAPIShPath.au3>
#include <StringConstants.au3>

#include "_Logging.au3"

Func _IsSafeApp(ByRef $sApp)

	Local $aApp
	Local $bSafe = False

	$sApp = StringStripWS($sApp, $STR_STRIPLEADING+$STR_STRIPTRAILING)
	$aApp = StringSplit($sApp, " ")

	For $iLoop = 1 To $aApp[0] Step 1
		If StringInStr($aApp[$iLoop], "=") Then
			Switch StringSplit($aApp[$iLoop], "=")[1]
				Case "--app-id"
					ContinueCase
				Case "--app-fallback-url"
					ContinueCase
				Case "--display-mode"
					ContinueCase
				Case "--ip-proc-id"
					ContinueCase
				Case "--mojo-named-platform-channel-pipe"
					ContinueCase
				Case "--ip-aumid"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		Else
			Switch $aApp[$iLoop]
				Case "--windows-store-app"
					ContinueCase
				Case "--ip-binding"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		EndIf
	Next

	If Not $bSafe Then FileWrite($hLogs[$AppSecurity], _NowCalc() & " - " & "Blocked Unsafe App: " & $sApp & @CRLF)

	Return $bSafe

EndFunc

Func _IsSafeURL(ByRef $sURL)

	Local $aURL
	Local $bSafe = False

	$aURL = StringSplit($sURL, ":")
	If $aURL[0] < 2 Then
		ReDim $aURL[3]
		$aURL[2] = $aURL[1]
		$aURL[1] = "https"
		$sURL = "https://" & $sURL
	EndIf

	Select
		Case $aURL[1] <> "http" And $aURL[1] <> "https"
			ContinueCase
		Case _WinAPI_UrlIs($sURL, $URLIS_FILEURL)
			ContinueCase
		Case _WinAPI_UrlIs($sURL, $URLIS_OPAQUE)
			$bSafe = False
		Case _WinAPI_UrlIs($sURL, $URLIS_URL)
			$bSafe = True
		Case Else
			;;;
	EndSelect

	If Not $bSafe Then FileWrite($hLogs[$AppSecurity], _NowCalc() & " - " & "Blocked Unsafe URL: " & $sURL & @CRLF)

	Return $bSafe

EndFunc
