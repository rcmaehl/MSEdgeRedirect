#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\MSEdgeRedirect.ico
#AutoIt3Wrapper_Outfile=MSEdgeRedirect_x86.exe
#AutoIt3Wrapper_Outfile_x64=MSEdgeRedirect.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=A Tool to Redirect News, Search, and Weather Results to Your Default Browser
#AutoIt3Wrapper_Res_Fileversion=0.2.1.0
#AutoIt3Wrapper_Res_ProductName=MSEdgeRedirect
#AutoIt3Wrapper_Res_ProductVersion=0.2.1.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 -v1 -v2 -v3
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <Misc.au3>
#include <Array.au3>
#include <String.au3>
#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>
#include <TrayConstants.au3>
#include <MsgBoxConstants.au3>

Global $hLogs[3] = _
	[FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppFailures.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppGeneral.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\URIFailures.log", $FO_APPEND)]

Global $aEdges[5] = [4, _
	"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
	"C:\Program Files (x86)\Microsoft\Edge Beta\Application\msedge.exe", _
	"C:\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe", _
	@LocalAppDataDir & "Microsoft\Edge SXS\Application\msedge.exe"]

Global $sVersion = "0.2.1.0"

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("GUICloseOnESC", 0)

If @OSArch = "X64" And _WinAPI_IsWow64Process() Then
	MsgBox($MB_ICONERROR+$MB_OK, "Wrong Version", "The 64-bit Version of MSEdgeRedirect must be used with 64-bit Windows!")
	FileWrite($hLogs[0], _NowCalc() & " - " & "32 Bit Version on 64 Bit System." & @CRLF)
	For $iLoop = 0 To UBound($hLogs) - 1
		FileClose($hLogs[$iLoop])
	Next
	Exit 1
EndIf

ProcessCMDLine()

Func ActiveMode(ByRef $aCMDLine)

	Local $sCMDLine = ""

	If RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER", "FilterFullPath") = $CmdLine[1] Then
		For $iLoop = 2 To $aCMDLine[0]
			$sCMDLine &= $aCMDLine[$iLoop] & " "
		Next
		_DecodeAndRun($sCMDLine)
		Exit
	Else
		MsgBox(0, "TEST - " & $aCMDLine[1], $sCMDLine)
		_DecodeAndRun($sCMDLine)
		Exit
	EndIf

EndFunc

Func IsInstalled()

	Local $sInstalledVer

	$sInstalledVer = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdge Redirect", "DisplayVersion")
	If @error Then
		$sInstalledVer = RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdge Redirect", "DisplayVersion")
		If @error Then
			RunSetup()
		EndIf
	EndIf
	If _VersionCompare($sVersion, $sInstalledVer) Then RunSetup(True)

EndFunc

Func ProcessCMDLine()

	Local $bHide = False
	Local $iParams = $CmdLine[0]

	If $iParams > 0 Then

		_ArrayDisplay($CmdLine)
		If _ArraySearch($aEdges, $CmdLine[1]) Then ; Image File Execution Options Mode
			ActiveMode($CmdLine)
		EndIf
		Do
			Switch $CmdLine[1]
				Case "/?", "/h", "/help"
					MsgBox(0, "Help and Flags", _
							"Checks PC for Windows 11 Release Compatibility" & @CRLF & _
							@CRLF & _
							"MSEdgeRedirect [/hide]" & @CRLF & _
							@CRLF & _
							@TAB & "/hide  " & @TAB & "Hides the tray icon" & @CRLF & _
							@TAB & "/update" & @TAB & "Downloads the latest RELEASE (default) or DEV build" & @CRLF & _
							@CRLF & _
							@CRLF)
					Exit 0
				Case "/h", "/hide"
					$bHide = True
					_ArrayDelete($CmdLine, 1)
				Case "/u", "/update"
					Select
						Case UBound($CmdLine) = 2
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, 1)
						Case UBound($CmdLine) > 2 And $CmdLine[2] = "dev"
							InetGet("https://nightly.link/rcmaehl/MSEdgeRedirect/workflows/mser/main/mser.zip", @ScriptDir & "\WhyNotWin11_dev.zip")
							_ArrayDelete($CmdLine, "1-2")
						Case UBound($CmdLine) > 2 And $CmdLine[2] = "release"
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, "1-2")
						Case StringLeft($CmdLine[2], 1) = "/"
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, 1)
						Case Else
							MsgBox(0, "Invalid", 'Invalid release type - "' & $CmdLine[2] & "." & @CRLF)
							Exit 87 ; ERROR_INVALID_PARAMETER
					EndSelect
				Case Else
					If @Compiled Then ; support for running non-compiled script - mLipok
						MsgBox(0, "Invalid", 'Invalid parameter - "' & $CmdLine[1] & "." & @CRLF)
						Exit 87 ; ERROR_INVALID_PARAMETER
					EndIf
			EndSwitch
		Until UBound($CmdLine) <= 1
	EndIf

	;IsInstalled()
	ReactiveMode($bHide)

EndFunc

Func ReactiveMode($bHide = False)

	Local $aMUI[2] = [Null, @MUILang]
	Local $aAdjust

	Local $hMsg

	Local $oError = ObjEvent("AutoIt.Error", "_LogError")
	#forceref $oError

	Local $Obj  = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	Local $hObj = ObjCreate("WbemScripting.SWbemSink")

	If IsObj($Obj) And IsObj($hObj) Then
		ObjEvent($hObj, "SINK_")
		$Obj.ExecNotificationQueryAsync($hObj, "SELECT * FROM __InstanceCreationEvent WITHIN 0.001 WHERE TargetInstance ISA 'Win32_Process'")
	EndIf
	#forceref SINK_OnObjectReady

	; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
	Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))

	_WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

	Local $hStartup = TrayCreateItem("Start With Windows")
	Local $hUpdate = TrayCreateItem("Check for Updates")
	TrayCreateItem("")
	Local $hDonate = TrayCreateItem("Donate")
	TrayCreateItem("")
	Local $hHide = TrayCreateItem("Hide Icon")
	Local $hExit = TrayCreateItem("Exit")

	If $bHide Then TraySetState($TRAY_ICONSTATE_HIDE)

	If FileExists(@StartupDir & "\MSEdgeRedirect.lnk") Then TrayItemSetState($hStartup, $TRAY_CHECKED)

	While True
		$hMsg = TrayGetMsg()

		Select

			Case $hMsg = $hHide
				TraySetState($TRAY_ICONSTATE_HIDE)

			Case $hMsg = $hExit
				ExitLoop

			Case $hMsg = $hDonate
				ShellExecute("https://paypal.me/rhsky")

			Case $hMsg = $hUpdate
				Switch _GetLatestRelease($sVersion)
					Case -1
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Test Build?"), _Translate($aMUI[1], "You're running a newer build than publicly Available!"), 10)
					Case 0
						Switch @error
							Case 0
								MsgBox($MB_OK + $MB_ICONINFORMATION + $MB_TOPMOST, _Translate($aMUI[1], "Up to Date"), _Translate($aMUI[1], "You're running the latest build!"), 10)
							Case 1
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Unable to Check for Updates"), _Translate($aMUI[1], "Unable to load release data."), 10)
							Case 2
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Unable to Check for Updates"), _Translate($aMUI[1], "Invalid Data Received!"), 10)
							Case 3
								Switch @extended
									Case 0
										MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Unable to Check for Updates"), _Translate($aMUI[1], "Invalid Release Tags Received!"), 10)
									Case 1
										MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Unable to Check for Updates"), _Translate($aMUI[1], "Invalid Release Types Received!"), 10)
								EndSwitch
						EndSwitch
					Case 1
						If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _Translate($aMUI[1], "Update Available"), _Translate($aMUI[1], "An Update is Available, would you like to download it?"), 10) = $IDYES Then ShellExecute("https://fcofix.org/MSEdgeRedirect/releases")
				EndSwitch

			Case $hMsg = $hStartup
				If Not FileExists(@StartupDir & "\MSEdgeRedirect.lnk") Then
					FileCreateShortcut(@AutoItExe, @StartupDir & "\MSEdgeRedirect.lnk")
					TrayItemSetState($hStartup, $TRAY_CHECKED)
				ElseIf FileExists(@StartupDir & "\MSEdgeRedirect.lnk") Then
					FileDelete(@StartupDir & "\MSEdgeRedirect.lnk")
					TrayItemSetState($hStartup, $TRAY_UNCHECKED)
				EndIf

			Case Else

		EndSelect
	WEnd

	_WinAPI_AdjustTokenPrivileges($hToken, $aAdjust, 0, $aAdjust)
	_WinAPI_CloseHandle($hToken)
	For $iLoop = 0 To UBound($hLogs) - 1
		FileClose($hLogs[$iLoop])
	Next
	Exit

EndFunc

Func RunSetup($bUpdate = False)
	#forceref $bUpdate

	Local $bIsAdmin = IsAdmin()

	If Not $bIsAdmin Then
		If MsgBox($MB_YESNO+$MB_ICONWARNING+$MB_DEFBUTTON2+$MB_SETFOREGROUND, _
			"Not Running As Admin", _
			"MSEdge Redirect is not installed and will now do so." & @CRLF & _
			@CRLF & _
			"However, you are not running as admin. Some features, such as ACTIVE MODE, will not be available to install." & @CRLF & _
			@CRLF & _
			"Continue?") = $IDNO Then Exit
	EndIf

	DirCreate(@LocalAppDataDir & "\MSEdgeRedirect\logs\")
	DirCreate(@LocalAppDataDir & "\MSEdgeRedirect\langs\")

EndFunc

Func SINK_OnObjectReady($OB)

	Local $Str, $Owner, $Ret
	#forceref $Ret

	Local $sCommandline


    Switch $OB.Path_.Class
		Case "__InstanceCreationEvent"
			If StringInStr($OB.TargetInstance.CommandLine, "microsoft-edge:") Then
				ProcessClose($OB.TargetInstance.ProcessID)
				$sCommandline = StringReplace($OB.TargetInstance.CommandLine, '"' & $ob.targetinstance.executablepath & '" ', "")
				If _ArraySearch($aEdges, $ob.targetinstance.executablepath, 1, $aEdges[0]) Then
					_DecodeAndRun($sCommandline)
				Else
					FileWrite($hLogs[2], _NowCalc() & " - Invalid MSEdge: " & $ob.targetinstance.executablepath & @CRLF)
				EndIf
			EndIf
            $str &= $OB.TargetInstance.ProcessID & "-"
            $str &= $ob.targetinstance.name & "-"
            $str &= $ob.targetinstance.csname & "-"
            $ret = $ob.targetinstance.getowner($owner)
            $str &= $ob.targetinstance.creationdate & "-"
            $str &= $ob.targetinstance.parentprocessid & "-"
            $str &= $ob.targetinstance.executablepath & @cr
            consolewrite("!->> Started  " & $str)
            $str = ""
    EndSwitch
    Return 1
EndFunc

Func SetupAppdata()
	Select
		Case Not FileExists(@LocalAppDataDir & "\MSEdgeRedirect\")
			DirCreate(@LocalAppDataDir & "\MSEdgeRedirect\logs\")
			ContinueCase
		Case Not FileExists(@LocalAppDataDir & "\MSEdgeRedirect\Langs\")
			DirCreate(@LocalAppDataDir & "\MSEdgeRedirect\langs\")
		Case Else
			;;;
	EndSelect
EndFunc

Func _DecodeAndRun($sCMDLine)

	Local $aLaunchContext

	Select
		Case StringRegExp($sCMDLine, "microsoft-edge:[\/]*?\?launchContext1")
			$aLaunchContext = StringSplit($sCMDLine, "=")
			If $aLaunchContext[0] >= 3 Then
				$sCMDLine = _UnicodeURLDecode($aLaunchContext[$aLaunchContext[0]])
				If _WinAPI_UrlIs($sCMDLine) Then
					ShellExecute($sCMDLine)
				Else
					FileWrite($hLogs[2], _NowCalc() & " - Invalid Regexed URL: " & $sCMDLine & @CRLF)
				EndIf
			EndIf
		Case Else
			$sCMDLine = StringRegExpReplace($sCMDLine, "--single-argument microsoft-edge:[\/]*", "")
			If _WinAPI_UrlIs($sCMDLine) Then
				ShellExecute($sCMDLine)
			Else
				FileWrite($hLogs[2], _NowCalc() & " - Invalid URL: " & $sCMDLine & @CRLF)
			EndIf
	EndSelect
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetLatestRelease
; Description ...: Checks GitHub for the Latest Release
; Syntax ........: _GetLatestRelease($sCurrent)
; Parameters ....: $sCurrent            - a string containing the current program version
; Return values .: Returns True if Update Available
; Author ........: rcmaehl
; Modified ......: 11/11/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetLatestRelease($sCurrent)

	Local $dAPIBin
	Local $sAPIJSON

	$dAPIBin = InetRead("https://api.fcofix.org/repos/rcmaehl/MSEdgeRedirect/releases")
	If @error Then Return SetError(1, 0, 0)
	$sAPIJSON = BinaryToString($dAPIBin)
	If @error Then Return SetError(2, 0, 0)

	Local $aReleases = _StringBetween($sAPIJSON, '"tag_name":"', '",')
	If @error Then Return SetError(3, 0, 0)
	Local $aRelTypes = _StringBetween($sAPIJSON, '"prerelease":', ',')
	If @error Then Return SetError(3, 1, 0)
	Local $aCombined[UBound($aReleases)][2]

	For $iLoop = 0 To UBound($aReleases) - 1 Step 1
		$aCombined[$iLoop][0] = $aReleases[$iLoop]
		$aCombined[$iLoop][1] = $aRelTypes[$iLoop]
	Next

	Return _VersionCompare($aCombined[0][0], $sCurrent)

EndFunc   ;==>_GetLatestRelease

Func _LogError($oError)
    FileWrite($hLogs[0], _
            "! ######################## OBJECT ERROR! #########################################" & @CRLF & _
            "!                err.number is        : " & @TAB & hex($oError.number,8) & @CRLF & _
            "!                err.scriptline is    : " & @TAB & $oError.scriptline & @CRLF & _
            "!                err.windesc is       : " & @TAB & $oError.windescription & @CRLF & _
            "!                err.desc is          : " & @TAB & $oError.description & @CRLF & _
            "!                err.source is        : " & @TAB & $oError.source & @CRLF & _
            "!                err.retcode is       : " & @TAB & hex($oError.retcode,8) & @CRLF & _
            "! ################################################################################" & @CRLF _
            )
    Return 0
EndFunc

;===============================================================================
; _UnicodeURLDecode()
; Description: : Tranlates a URL-friendly string to a normal string
; Parameter(s): : $toDecode - The URL-friendly string to decode
; Return Value(s): : The URL decoded string
; Author(s): : nfwu, Dhilip89
; Note(s): : Modified from _URLDecode() that's only support non-unicode.
;
;===============================================================================
Func _UnicodeURLDecode($toDecode)
    Local $strChar = "", $iOne, $iTwo
    Local $aryHex = StringSplit($toDecode, "")
    For $i = 1 To $aryHex[0]
        If $aryHex[$i] = "%" Then
            $i = $i + 1
            $iOne = $aryHex[$i]
            $i = $i + 1
            $iTwo = $aryHex[$i]
            $strChar = $strChar & Chr(Dec($iOne & $iTwo))
        Else
            $strChar = $strChar & $aryHex[$i]
        EndIf
    Next
    Local $Process = StringToBinary (StringReplace($strChar, "+", " "))
    Local $DecodedString = BinaryToString ($Process, 4)
    Return $DecodedString
EndFunc   ;==>_UnicodeURLDecode

#Region Translation Functions
Func _GetFile($sFile, $sFormat = $FO_READ)
	Local Const $hFileOpen = FileOpen($sFile, $sFormat)
	If $hFileOpen = -1 Then
		Return SetError(1, 0, '')
	EndIf
	Local Const $sData = FileRead($hFileOpen)
	FileClose($hFileOpen)
	Return $sData
EndFunc   ;==>_GetFile

Func _INIUnicode($sINI)
	If FileExists($sINI) = 0 Then
		Return FileClose(FileOpen($sINI, $FO_OVERWRITE + $FO_UNICODE))
	Else
		Local Const $iEncoding = FileGetEncoding($sINI)
		Local $fReturn = True
		If Not ($iEncoding = $FO_UNICODE) Then
			Local $sData = _GetFile($sINI, $iEncoding)
			If @error Then
				$fReturn = False
			EndIf
			_SetFile($sData, $sINI, $FO_APPEND + $FO_UNICODE)
		EndIf
		Return $fReturn
	EndIf
EndFunc   ;==>_INIUnicode

Func _SetFile($sString, $sFile, $iOverwrite = $FO_READ)
	Local Const $hFileOpen = FileOpen($sFile, $iOverwrite + $FO_APPEND)
	FileWrite($hFileOpen, $sString)
	FileClose($hFileOpen)
	If @error Then
		Return SetError(1, 0, False)
	EndIf
	Return True
EndFunc   ;==>_SetFile

Func _Translate($iMUI, $sString)
	Local $sReturn
	_INIUnicode(@LocalAppDataDir & "\MSEdgeRedirect\Langs\" & $iMUI & ".lang")
	$sReturn = IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Langs\" & $iMUI & ".lang", "Strings", $sString, $sString)
	$sReturn = StringReplace($sReturn, "\n", @CRLF)
	Return $sReturn
EndFunc   ;==>_Translate
#EndRegion