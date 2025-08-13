#include-once

#include <Array.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>
#include <StringConstants.au3>

#include "_Logging.au3"

; URI Scheme Regex
; ^[a-zA-Z][a-zA-Z0-9+\-.]*$:

Func _ArraySafeDelim($aArray)
    
    Local $sChars = _ArrayToString($aArray, "")
    
    Select
        
        Case Not StringInStr($sChars, "|")
            Return "|"
        Case Else
            For $iLoop = 33 To 191 Step 1
                Switch $iLoop
                    Case 48 to 57 ; Skip Numbers
                        ContinueLoop
                    Case 65 to 90 ; Skip Upper Case
                        ContinueLoop
                    Case 97 to 122 ; Skip Lower Case
                        ContinueLoop
                    Case Else
                        If Not StringInStr($sChars, Chr($iLoop)) Then Return Chr($iLoop)
                EndSwitch
            Next
            
    EndSelect

    ; If No Char Found, Error
    Return SetError(1, 0, False)

EndFunc

Func _IsSafeApp(ByRef $sApp)

	Local $aApp
	Local $bSafe = False

	$sApp = StringStripWS($sApp, $STR_STRIPLEADING+$STR_STRIPTRAILING)
	$aApp = StringSplit($sApp, " ")

	For $iLoop = 1 To $aApp[0] Step 1
		If StringInStr($aApp[$iLoop], "=") Then
			Switch StringSplit($aApp[$iLoop], "=")[1]
				Case "--app-fallback-url"
					ContinueCase
				Case "--app-id"
					ContinueCase
				Case "--display-mode"
					ContinueCase
				Case "--ip-aumid"
					ContinueCase
				Case "--ip-proc-id"
					ContinueCase
				Case "--mojo-named-platform-channel-pipe"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		Else
			Switch $aApp[$iLoop]
				Case "--ip-binding"
					ContinueCase
				Case "--windows-store-app"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		EndIf
	Next

	If Not $bSafe Then _Log($hLogs[$AppSecurity], "Blocked Unsafe App: " & $sApp & @CRLF)

	Return $bSafe

EndFunc

Func _IsSafeFile(ByRef $sPath)

	Local $aPath
	Local $bSafe = False

	If Not FileExists($sPath) Then $bSafe = False

	$aPath = StringSplit($sPath, ".")	

	Switch $aPath[$aPath[0]]
		Case "htm", "html", "xhtml" ; HTML
			$bSafe = True
		Case "xml" ; XML
			$bSafe = True
		Case "mht", "mhtml" ; Archives
			$bSafe = True
		Case "jpg", "jpeg", "png", "apng", "gif", "png", "webp", "svg", "bmp", "ico", "avif" ; Images
			$bSafe = True
		Case "pdf", "txt", "csv" ; Documents
			$bSafe = True
		Case "mp4", "webm", "ogg", "mp3", "wav", "m4a" ; Media
			$bSafe = True
		Case "rss", "atom" ; Feeds
			$bSafe = True
		Case Else
			$bSafe = False
	EndSwitch

	If Not $bSafe Then _Log($hLogs[$AppSecurity], "Blocked Potentially Unsafe File: " & $sPath & @CRLF)

	Return $bSafe

EndFunc

Func _IsSafeFlag(ByRef $sCommandLine)

	Local $aCMDLine
	Local $bSafe = False

	$sCommandLine = StringStripWS($sCommandLine, $STR_STRIPLEADING+$STR_STRIPTRAILING)
	$aCMDLine = StringSplit($sCommandLine, " ")

	For $iLoop = 1 To $aCMDLine[0] Step 1
		If StringInStr($aCMDLine[$iLoop], "=") Then
			Switch StringSplit($aCMDLine[$iLoop], "=")[1]
				Case "--app-fallback-url"
					ContinueCase
				Case "--app-id"
					ContinueCase
				Case "--autoplay-policy"
					ContinueCase
				Case "--display-mode"
					ContinueCase
				Case "--ip-aumid"
					ContinueCase
				Case "--ip-proc-id"
					ContinueCase
				Case "--mojo-named-platform-channel-pipe"
					ContinueCase
				Case "--profile-directory"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		Else
			Switch $aCMDLine[$iLoop]
				Case "--from-installer"
					ContinueCase
				Case "--inprivate"
					ContinueCase
				Case "--ip-binding"
					ContinueCase
				Case "--kiosk"
					ContinueCase
				Case "--suspend-background-mode"
					ContinueCase
				Case "--windows-store-app"
					ContinueCase
				Case "--winrt-background-task-event"
					$bSafe = True
				Case Else
					$bSafe = False
					ExitLoop
			EndSwitch
		EndIf
	Next

	If Not $bSafe Then _Log($hLogs[$AppSecurity], "Blocked Unsafe Flag: " & $sCommandLine & @CRLF)

	Return $bSafe

EndFunc

#cs
Func _IsSafePDF(ByRef $sPath)

	Local $bSafe = False

	Select
		Case Not FileExists($sPath)
			$bSafe = False
		Case Not (StringRight($sPath, 4) = ".pdf")
			$bSafe = False
		Case Else
			$bSafe = True
	EndSelect

	If Not $bSafe Then _Log($hLogs[$AppSecurity], "Blocked Unsafe File: " & $sPath & @CRLF)

	Return $bSafe

EndFunc
#ce

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

	If Not $bSafe Then _Log($hLogs[$AppSecurity], "Blocked Unsafe URL: " & $sURL & @CRLF)

	Return $bSafe

EndFunc

Func _SafeRun($sPath, $sCMDLine = "")

	Local $tProcess = DllStructCreate($tagPROCESS_INFORMATION)
	Local $tStartup = DllStructCreate($tagSTARTUPINFO)	

	If $sCMDLine = "" Then
		_WinAPI_CreateProcess('', $sPath, 0, 0, 0, $CREATE_NEW_PROCESS_GROUP, 0, 0, $tStartup, $tProcess)
	Else
		_WinAPI_CreateProcess($sPath, "--MSEdgeRedirect " & $sCMDLine, 0, 0, 0, $CREATE_NEW_PROCESS_GROUP, 0, 0, $tStartup, $tProcess)
	EndIf

EndFunc