#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\MSEdgeRedirect.ico
#AutoIt3Wrapper_Outfile=MSEdgeRedirect_x86.exe
#AutoIt3Wrapper_Outfile_x64=MSEdgeRedirect.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=https://www.msedgeredirect.com
#AutoIt3Wrapper_Res_Description=A Tool to Redirect News, Search, Widgets, Weather and More to Your Default Browser
#AutoIt3Wrapper_Res_Fileversion=0.5.0.0
#AutoIt3Wrapper_Res_ProductName=MSEdgeRedirect
#AutoIt3Wrapper_Res_ProductVersion=0.5.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Compatibility=Win8,Win81,Win10
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 -v1 -v2 -v3
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so
#AutoIt3Wrapper_Res_Icon_Add=Assets\MSEdgeRedirect.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <Misc.au3>
#include <Array.au3>
#include <String.au3>
#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>
#include <EditConstants.au3>
#include <TrayConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>

#include "Includes\_Theming.au3"

#include "Includes\ResourcesEx.au3"

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("GUICloseOnESC", 0)

Global $aEdges[5] = [4, _
	"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
	"C:\Program Files (x86)\Microsoft\Edge Beta\Application\msedge.exe", _
	"C:\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe", _
	@LocalAppDataDir & "\Microsoft\Edge SXS\Application\msedge.exe"]

Global $sVersion = "0.5.0.0"

SetupAppdata()

Global $hLogs[3] = _
	[FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppFailures.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\AppGeneral.log", $FO_APPEND), _
	FileOpen(@LocalAppDataDir & "\MSEdgeRedirect\logs\URIFailures.log", $FO_APPEND)]

RunArchCheck()
RunHTTPCheck()
ProcessCMDLine()

Func ActiveMode(ByRef $aCMDLine)

	Local $sCMDLine = ""

	Select
		Case $aCMDLine[0] = 1 ; No Parameters
			ReDim $aCMDLine[3]
			$aCMDLine[2] = ""
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--inprivate" ; In Private Browsing, No Parameters
			$aCMDLine[1] = StringReplace($aCMDLine[1], "msedge.exe", "msedge_no_ifeo.exe")
			ShellExecute($aCMDLine[1], $aCMDLine[2])
		Case Else
			For $iLoop = 2 To $aCMDLine[0]
				$sCMDLine &= $aCMDLine[$iLoop] & " "
			Next
			_DecodeAndRun($sCMDLine)
	EndSelect

EndFunc

Func ProcessCMDLine()

	Local $aMUI[2] = [Null, @MUILang]
	Local $aPIDs
	Local $bHide = False
	Local $iParams = $CmdLine[0]
	Local $bSilent = False
	Local $aInstall[3]
	Local $bPortable = False

	If DriveGetType(@ScriptDir) = "Removable" Then $bPortable = True

	If $iParams > 0 Then

		;_ArrayDisplay($CmdLine)
		If _ArraySearch($aEdges, $CmdLine[1]) > 0 Then ; Image File Execution Options Mode
			ActiveMode($CmdLine)
			If _GetSettingValue("NoUpdates") And Random(1, 10, 1) = 1 Then
				Switch _GetLatestRelease($sVersion)
					Case -1
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Test Build?"), _Translate($aMUI[1], "You're running a newer build than publicly Available!"), 10)
					Case 1
						If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _Translate($aMUI[1], "Update Available"), _Translate($aMUI[1], "An Update is Available, would you like to download it?"), 10) = $IDYES Then ShellExecute("https://fcofix.org/MSEdgeRedirect/releases")
				EndSwitch
			EndIf
			Exit
		EndIf

		Do
			Switch $CmdLine[1]
				Case "/?", "/help"
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
				Case "/change"
					RunSetup(True, $bSilent)
					Exit
				Case "/h", "/hide"
					$bHide = True
					_ArrayDelete($CmdLine, 1)
				Case "/p", "/portable"
					$bPortable = True
					_ArrayDelete($CmdLine, 1)
				Case "/repair"
					RunRepair()
					Exit
				Case "/si", "/silentinstall"
					$bSilent = True
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
				Case "/uninstall"
					RunRemoval()
					Exit
				Case Else
					If @Compiled Then ; support for running non-compiled script - mLipok
						MsgBox(0, "Invalid", 'Invalid parameter - "' & $CmdLine[1] & "." & @CRLF)
						Exit 87 ; ERROR_INVALID_PARAMETER
					EndIf
			EndSwitch
		Until UBound($CmdLine) <= 1
	Else
		;;;
	EndIf

	If Not $bPortable Then
		$aInstall = _IsInstalled()

		Select
			Case Not $aInstall[0] ; Not Installed
				RunSetup(False, $bSilent)
			Case _VersionCompare($sVersion, $aInstall[2]) ; Installed, Out of Date
				RunSetup($aInstall[1], $bSilent)
			Case StringInStr($aInstall[1], "HKCU") ; Installed, Up to Date, Service Mode
				If Not @ScriptDir = @LocalAppDataDir & "\MSEdgeRedirect" Then
					ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", "", @LocalAppDataDir & "\MSEdgeRedirect\")
				Else
					$aPIDs = ProcessList(@ScriptName)
					For $iLoop = 1 To $aPIDs[0][0] Step 1
						If $aPIDs[$iLoop][1] <> @AutoItPID Then
							$bHide = False
							ProcessClose($aPIDs[$iLoop][1])
						EndIf
					Next
				EndIf
			Case Else
				Exit
		EndSelect
	EndIf
	ReactiveMode($bHide)

EndFunc

Func ReactiveMode($bHide = False)

	Local $aMUI[2] = [Null, @MUILang]
	Local $hTimer = TimerInit()
	Local $aAdjust

	Local $hMsg

	; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
	Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))

	_WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

	TrayCreateItem($sVersion)
	TrayItemSetState(-1, $TRAY_DISABLE)
	TrayCreateItem("")
	Local $hStartup = TrayCreateItem("Start With Windows")
	Local $hUpdate = TrayCreateItem("Check for Updates")
	TrayCreateItem("")
	Local $hDonate = TrayCreateItem("Donate")
	TrayCreateItem("")
	Local $hHide = TrayCreateItem("Hide Icon")
	Local $hExit = TrayCreateItem("Exit")

	If $bHide Then TraySetState($TRAY_ICONSTATE_HIDE)

	If FileExists(@StartupDir & "\MSEdgeRedirect.lnk") Then TrayItemSetState($hStartup, $TRAY_CHECKED)


	Local $aProcessList
	Local $sCommandline

	While True
		$hMsg = TrayGetMsg()

		If TimerDiff($hTimer) >= 100 Then
			$aProcessList = ProcessList("msedge.exe")
			For $iLoop = 1 To $aProcessList[0][0] - 1
				$sCommandline = _WinAPI_GetProcessCommandLine($aProcessList[$iLoop][1])
				If (StringInStr($sCommandline, "microsoft-edge:") And Not StringInStr($sCommandline, "--inprivate")) Or StringInStr($sCommandline, "--app-id") Then
					ProcessClose($aProcessList[$iLoop][1])
					If _ArraySearch($aEdges, _WinAPI_GetProcessFileName($aProcessList[$iLoop][1]), 1, $aEdges[0]) > 0 Then
						_DecodeAndRun($sCommandline)
					EndIf
				EndIf
			Next
			$hTimer = TimerInit()
		EndIf

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
					FileCreateShortcut(@AutoItExe, @StartupDir & "\MSEdgeRedirect.lnk", @ScriptDir)
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

Func RunArchCheck()
	If @Compiled And @OSArch = "X64" And _WinAPI_IsWow64Process() Then
		MsgBox($MB_ICONERROR+$MB_OK, "Wrong Version", "The 64-bit Version of MSEdgeRedirect must be used with 64-bit Windows!")
		FileWrite($hLogs[0], _NowCalc() & " - " & "32 Bit Version on 64 Bit System. EXITING!" & @CRLF)
		For $iLoop = 0 To UBound($hLogs) - 1
			FileClose($hLogs[$iLoop])
		Next
		Exit 216 ; ERROR_EXE_MACHINE_TYPE_MISMATCH
	EndIf
EndFunc

Func RunHTTPCheck()

	Local $sHive = ""

	If _WinAPI_IsWow64Process() Then
		$sHive = "HKCU64"
	Else
		$sHive = "HKCU"
	EndIf

	If StringInStr(RegRead($sHive & "\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", "ProgId"), "MSEdge") Or _
		StringInStr(RegRead($sHive & "\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice", "ProgId"), "MSEdge") Then
		MsgBox($MB_ICONERROR+$MB_OK, "Edge Set As Default", "You must set a different Default Browser to use MSEdgeRedirect!")
		FileWrite($hLogs[0], _NowCalc() & " - " & "Found MS Edge set as default browser, EXITING!" & @CRLF)
		For $iLoop = 0 To UBound($hLogs) - 1
			FileClose($hLogs[$iLoop])
		Next
		Exit 4315 ; ERROR_MEDIA_INCOMPATIBLE
	EndIf

EndFunc

Func RunInstall(ByRef $aConfig, ByRef $aSettings)

	Local $sArgs = ""
	Local Enum $bManaged, $vMode
	Local Enum $bNoApps, $bNoBing, $bNoPDFs, $bNoTray, $bNoUpdates, $sPDFApp, $sSearch, $sSearchPath, $sStartMenu, $bStartup

	SetOptionsRegistry("NoApps", $aSettings[$bNoApps], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoBing", $aSettings[$bNoBing], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoPDFs", $aSettings[$bNoPDFs], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoTray", $aSettings[$bNoTray], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoUpdates", $aSettings[$bNoUpdates], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("PDFApp", $aSettings[$sPDFApp], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("Search", $aSettings[$sSearch], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("SearchPath", $aSettings[$sSearchPath], $aConfig[$vMode], $aConfig[$bManaged])

	If $aConfig[$vMode] Then
		FileCopy(@ScriptFullPath, "C:\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE)
	Else
		If $aSettings[$bNoTray] Then $sArgs = "/hide"
		FileCopy(@ScriptFullPath, @LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE)
		If $aSettings[$bStartup] Then FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @StartupDir & "\MSEdgeRedirect.lnk")
		Switch $aSettings[$sStartMenu]

			Case "Full"
				DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
				FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\Settings.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", "/change")
				ContinueCase

			Case "App Only"
				DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
				FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\MSEdgeRedirect.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", $sArgs)

			Case Else
				;;;

		EndSwitch
	EndIf
EndFunc

Func RunRemoval($bUpdate = False)

	Local $aPIDs
	Local $sHive = ""
	Local $sLocation = ""

	$aPIDs = ProcessList("msedgeredirect.exe")
	For $iLoop = 1 To $aPIDs[0][0] Step 1
		If $aPIDs[$iLoop][1] <> @AutoItPID Then ProcessClose($aPIDs[$iLoop][1])
	Next

	If IsAdmin() Then
		$sLocation = "C:\Program Files\MSEdgeRedirect\"
		If _WinAPI_IsWow64Process() Then
			$sHive = "HKLM64"
		Else
			$sHive = "HKLM"
		EndIf
	Else
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
		If _WinAPI_IsWow64Process() Then
			$sHive = "HKCU64"
		Else
			$sHive = "HKCU"
		EndIf
	EndIf

	; App Paths
	RegDelete($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe")

	; App Settings
	RegDelete($sHive & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect")

	; URI Handler for Pre Win11 22494 Installs
	RegDelete($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge")

	; Generic Program Info
	RegDelete($sHive & "\Software\Classes\MSEdgeRedirect")
	RegDelete($sHive & "\Software\Classes\Applications\MSEdgeRedirect.exe")

	; IFEO
	RegDelete($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe")

	; Uninstall Info
	RegDelete($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect")

	FileDelete(@StartupDir & "\MSEdgeRedirect.lnk")
	DirRemove(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect", $DIR_REMOVE)

	If IsAdmin() Then
		For $iLoop = 1 To $aEdges[0] Step 1
			If FileExists(StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe")) Then
				FileDelete(StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe"))
			EndIf
		Next
	EndIf

	If $bUpdate Then
		FileDelete($sLocation & "*")
	Else
		Run(@ComSpec & " /c " & 'ping google.com && del /Q "' & $sLocation & '*"', "", @SW_HIDE)
		Exit
	EndIf

EndFunc

Func RunRepair()

	If IsAdmin() Then
		For $iLoop = 1 To $aEdges[0] Step 1
			If FileExists(StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe")) Then
				FileCopy($aEdges[$iLoop], StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe"), $FC_OVERWRITE)
			EndIf
		Next
		Exit
	Else
		Exit 5 ; ERROR_ACCESS_DENIED
	EndIf

EndFunc

Func RunSetup($bUpdate = False, $bSilent = False)
	#forceref $bSilent

	Local $aMUI[2] = [Null, @MUILang]
	Local $hMsg
	Local $sArgs = ""
	Local $sEdges
	Local $sEngine
	Local $aHandler
	Local $sHandler
	Local $bIsAdmin = IsAdmin()
	Local $hChannels[4]
	Local $aChannels[4] = [True, True, False, False]

	Local $aConfig[2] = [False, "Service"] ; Default Setup.ini Values
	Local Enum $bManaged, $vMode

	Local $aSettings[10] = [False, False, False, False, False, "", "", "", "Full", True]
	Local Enum $bNoApps, $bNoBing, $bNoPDFs, $bNoTray, $bNoUpdates, $sPDFApp, $sSearch, $sSearchPath, $sStartMenu, $bStartup

	If $bSilent Then

		If Not FileExists(@ScriptDir & "\Setup.ini") Then Exit 2 ; ERROR_FILE_NOT_FOUND

		$aConfig[$bManaged] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Config", "Managed", False))
		$aConfig[$vMode] = IniRead(@ScriptDir & "\Setup.ini", "Config", "Mode", "Service")

		If $aConfig[$vMode] = "active" Then
			$aConfig[$vMode] = True
		Else
			$aConfig[$vMode] = False
		EndIf

		If ($aConfig[$bManaged] Or $aConfig[$vMode]) And Not IsAdmin() Then Exit 5 ; ERROR_ACCESS_DENIED

		$sEdges = IniRead(@ScriptDir & "\Setup.ini", "Settings", "Edges", "")
		If StringInStr($sEdges, "Stable") Then $aChannels[0] = True
		If StringInStr($sEdges, "Beta") Then $aChannels[1] = True
		If StringInStr($sEdges, "Dev") Then $aChannels[2] = True
		If StringInStr($sEdges, "Canary") Then $aChannels[3] = True

		For $iLoop = 0 To 3 Step 1
			If $aChannels[$iLoop] = True Then ExitLoop
			If $iLoop = 3 Then Exit 160 ; ERROR_BAD_ARGUMENTS
		Next

		$aSettings[$bNoApps] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "NoApps", $aSettings[$bNoApps]))
		$aSettings[$bNoBing] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "NoBing", $aSettings[$bNoBing]))
		$aSettings[$bNoPDFs] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "NoPDFs", $aSettings[$bNoPDFs]))
		$aSettings[$bNoTray] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "NoTray", $aSettings[$bNoTray]))
		$aSettings[$bNoUpdates] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "NoUpdates", $aSettings[$bNoUpdates]))
		$aSettings[$sPDFApp] = IniRead(@ScriptDir & "\Setup.ini", "Settings", "PDFApp", $aSettings[$sPDFApp])
		$aSettings[$sSearch] = IniRead(@ScriptDir & "\Setup.ini", "Settings", "Search", $aSettings[$sSearch])
		$aSettings[$sSearchPath] = IniRead(@ScriptDir & "\Setup.ini", "Settings", "SearchPath", $aSettings[$sSearchPath])
		$aSettings[$sStartMenu] = IniRead(@ScriptDir & "\Setup.ini", "Settings", "StartMenu", $aSettings[$sStartMenu])
		$aSettings[$bStartup] = _Bool(IniRead(@ScriptDir & "\Setup.ini", "Settings", "Startup", $aSettings[$bStartup]))

		For $iLoop = $bNoApps To $bNoUpdates Step 1
			If Not IsBool($aSettings[$iLoop]) Then Exit 160 ; ERROR_BAD_ARGUMENTS
		Next
		If Not IsBool($aSettings[$bStartup]) Then Exit 160 ; ERROR_BAD_ARGUMENTS

		RunInstall($aConfig, $aSettings)
		SetAppRegistry($aConfig[$vMode])
		If $aConfig[$vMode] Then
			SetIFEORegistry($aChannels)
		Else
			If $aSettings[$bNoTray] Then $sArgs = "/hide"
			ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $sArgs, @LocalAppDataDir & "\MSEdgeRedirect\")
		EndIf
		Exit

	Else

		If _GetSettingValue("NoUpdates") Then
			Switch _GetLatestRelease($sVersion)
				Case -1
					MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _Translate($aMUI[1], "Test Build?"), _Translate($aMUI[1], "You're running a newer build than publicly Available!"), 10)
				Case 1
					If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _Translate($aMUI[1], "Update Available"), _Translate($aMUI[1], "An Update is Available, would you like to download it?"), 10) = $IDYES Then ShellExecute("https://fcofix.org/MSEdgeRedirect/releases")
			EndSwitch
		EndIf

		If StringInStr($bUpdate, "HKLM") And Not $bIsAdmin And Not @Compiled Then
			MsgBox($MB_ICONERROR+$MB_OK, "Admin Required", "Unable to update an Admin Install without Admin Rights!")
			FileWrite($hLogs[0], _NowCalc() & " - " & "Non Admin Update Attempt on Admin Install. EXITING!" & @CRLF)
			For $iLoop = 0 To UBound($hLogs) - 1
				FileClose($hLogs[$iLoop])
			Next
			Exit 5 ; ERROR_ACCESS_DENIED
		EndIf

		; Disable Scaling
		If @OSVersion = 'WIN_10' Then DllCall(@SystemDir & "\User32.dll", "bool", "SetProcessDpiAwarenessContext", "HWND", "DPI_AWARENESS_CONTEXT" - 1)

		Local $hInstallGUI = GUICreate("MSEdge Redirect " & $sVersion & " Setup", 640, 480)

		GUICtrlCreateLabel("", 0, 0, 180, 480)
		GUICtrlSetBkColor(-1, 0x00A4EF)

		GUICtrlCreateIcon("", -1, 26, 26, 128, 128)
		If @Compiled Then
			_SetBkSelfIcon(-1, "", 0x00A4EF, @ScriptFullPath, 201, 128, 128)
		Else
			_SetBkIcon(-1, "", 0x00A4EF, @ScriptDir & "\assets\MSEdgeRedirect.ico", -1, 128, 128)
		EndIf

		#Region License Page
		Local $hLicense = GUICreate("", 460, 480, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		FileInstall("./LICENSE", @LocalAppDataDir & "\MSEdgeRedirect\License.txt")

		If $bUpdate Then
			GUICtrlCreateLabel("Pleae read the following License. You must accept the terms of the license before continuing with the upgrade.", 20, 20, 420, 40)
		Else
			GUICtrlCreateLabel("Pleae read the following License. You must accept the terms of the license before continuing with the installation.", 20, 20, 420, 40)
		EndIf

		GUICtrlCreateEdit("TL;DR: It's FOSS, you can edit it, repackage it, eat it (not recommended), or throw it at your neighbor Steve (depends on the Steve), but changes to it must be LGPL v3 too." & _
			@CRLF & @CRLF & _
			FileRead(@LocalAppDataDir & "\MSEdgeRedirect\License.txt"), 20, 60, 420, 280, $ES_READONLY + $WS_VSCROLL)

		Local $hAgree = GUICtrlCreateRadio("I accept this license", 20, 350, 420, 20)
		Local $hDisagree = GUICtrlCreateRadio("I don't accept this license", 20, 370, 420, 20)
		GUICtrlSetState(-1, $GUI_CHECKED)

		Local $hNext = GUICtrlCreateButton("NEXT", 20, 410, 420, 50)
		GUICtrlSetState(-1, $GUI_DISABLE)

		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Settings Page
		Local $hSettings = GUICreate("", 460, 480, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		If $bUpdate Then
			GUICtrlCreateLabel("MSEdge Redirect " & $sVersion & " Update", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Install to update MS Edge Redirect after customizing your preferred options", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Install MSEdge Redirect " & $sVersion, 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Install to install MS Edge Redirect after customizing your preferred options", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		GUICtrlCreateGroup("Mode", 20, 80, 420, 220)
			Local $hService = GUICtrlCreateRadio("Service Mode - Per User" & @CRLF & _
				@CRLF & _
				"MSEdge Redirect stays running in the background. Detected Edge data is redirected to your default browser.", _
				50, 100, 380, 60, $BS_TOP+$BS_MULTILINE)
			GUICtrlSetState(-1, $GUI_CHECKED)

			Local $hStartup = GUICtrlCreateCheckbox("Start MSEdge Redirect Service With Windows", 70, 160, 320, 20)
			Local $hNoIcon = GUICtrlCreateCheckbox("Hide MSEdge Redirect Service Icon from Tray", 70, 180, 320, 20)

			GUICtrlCreateIcon("imageres.dll", 78, 30, 210, 16, 16)
			Local $hActive = GUICtrlCreateRadio("Active Mode - All Users" & @CRLF & _
				@CRLF & _
				"MSEdge Redirect only runs when a selected Edge is launched, similary to the old EdgeDeflector app.", _
				50, 210, 380, 60, $BS_TOP+$BS_MULTILINE)

			$hChannels[0] = GUICtrlCreateCheckbox("Edge Stable", 70, 270, 90, 20)
			GUICtrlSetState(-1, $GUI_CHECKED)
			$hChannels[1] = GUICtrlCreateCheckbox("Edge Beta", 160, 270, 90, 20)
			$hChannels[2] = GUICtrlCreateCheckbox("Edge Dev", 250, 270, 90, 20)
			$hChannels[3] = GUICtrlCreateCheckbox("Edge Canary", 340, 270, 90, 20)

			GUICtrlSetState($hChannels[0], $GUI_DISABLE)
			GUICtrlSetState($hChannels[1], $GUI_DISABLE)
			GUICtrlSetState($hChannels[2], $GUI_DISABLE)
			GUICtrlSetState($hChannels[3], $GUI_DISABLE)

		GUICtrlCreateGroup("Options", 20, 300, 420, 100)
			Local $hNoApps = GUICtrlCreateCheckbox("De-embed Windows Store 'Apps'", 50, 320, 380, 20)
			Local $hNoPDFs = GUICtrlCreateCheckbox("Redirect PDFs to:", 50, 340, 240, 20)
			Local $hPDFPath = GUICtrlCreateLabel("",290, 340, 140, 20)
			Local $hSearch = GUICtrlCreateCheckbox("Replace Bing Search Results with:", 50, 360, 240, 20)
			Local $hEngine = GUICtrlCreateCombo("", 290, 355, 140, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "Ask|Baidu|Custom|DuckDuckGo|Ecosia|Google|Sogou|Yahoo|Yandex", "Google")
			GUICtrlSetState(-1, $GUI_DISABLE)


		Local $hInstall = GUICtrlCreateButton("Install", 20, 410, 420, 50)
		GUICtrlSetFont(-1, 16, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Finish Page
		Local $hFinish = GUICreate("", 460, 480, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		If $bUpdate Then
			GUICtrlCreateLabel("Updated Successfully", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Installed Successfully", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		Local $hClose = GUICtrlCreateButton("Close", 20, 410, 420, 50)
		GUICtrlSetFont(-1, 16, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)

		GUISwitch($hInstallGUI)
		#EndRegion

		GUISetState(@SW_SHOW, $hInstallGUI)
		GUISetState(@SW_SHOW, $hLicense)

		While True
			$hMsg = GUIGetMsg()

			Select

				Case $hMsg = $GUI_EVENT_CLOSE or $hMsg = $hClose
					If Not $aConfig[$vMode] Then
						If $aSettings[$bNoTray] Then $sArgs = "/hide"
						ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $sArgs, @LocalAppDataDir & "\MSEdgeRedirect\")
					EndIf
					Exit

				Case $hMsg = $hAgree or $hMsg = $hDisagree
					If _IsChecked($hAgree) Then
						GUICtrlSetState($hNext, $GUI_ENABLE)
					Else
						GUICtrlSetState($hNext, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hNext
					GUISetState(@SW_HIDE, $hLicense)
					GUISetState(@SW_SHOW, $hSettings)

				Case $hMsg = $hActive or $hMsg = $hService
					If _IsChecked($hActive) And Not IsAdmin() Then
						If ShellExecute(@ScriptFullPath, "", @ScriptDir, "RunAs") Then Exit
					EndIf
					If _IsChecked($hService) Then
						GUICtrlSetState($hInstall, $GUI_ENABLE)
						GUICtrlSetState($hStartup, $GUI_ENABLE)
						GUICtrlSetState($hNoIcon, $GUI_ENABLE)
						GUICtrlSetState($hChannels[0], $GUI_DISABLE)
						GUICtrlSetState($hChannels[1], $GUI_DISABLE)
						GUICtrlSetState($hChannels[2], $GUI_DISABLE)
						GUICtrlSetState($hChannels[3], $GUI_DISABLE)
					Else
						GUICtrlSetState($hStartup, $GUI_DISABLE)
						GUICtrlSetState($hNoIcon, $GUI_DISABLE)
						GUICtrlSetState($hChannels[0], $GUI_ENABLE)
						GUICtrlSetState($hChannels[1], $GUI_ENABLE)
						GUICtrlSetState($hChannels[2], $GUI_ENABLE)
						GUICtrlSetState($hChannels[3], $GUI_ENABLE)
						ContinueCase
					EndIf

				Case $hMsg = $hChannels[0] Or $hMsg = $hChannels[1] Or $hMsg = $hChannels[2] Or $hMsg = $hChannels[3]
					GUICtrlSetState($hInstall, $GUI_DISABLE)
					For $iLoop = 0 To 3 Step 1
						If _IsChecked($hChannels[$iLoop]) Then
							GUICtrlSetState($hInstall, $GUI_ENABLE)
							ExitLoop
						EndIf
					Next

				Case $hMsg = $hSearch
					If _IsChecked($hSearch) Then
						GUICtrlSetState($hEngine, $GUI_ENABLE)
					Else
						GUICtrlSetState($hEngine, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hEngine And GUICtrlRead($hEngine) = "Custom"
					$sEngine = InputBox("Enter Search Engine URL", "Enter the URL format of the custom search Engine to use", "https://duckduckgo.com/?q=")
					If @error Or Not _WinAPI_UrlIs($sEngine) Then GUICtrlSetData($hEngine, "Google")

				Case $hMsg = $hNoPDFs
					If _IsChecked($hNoPDFs) Then
						$sHandler = FileOpenDialog("Select a PDF Handler", @ProgramFilesDir, "Executables (*.exe)", $FD_FILEMUSTEXIST)
						If @error Then
							GUICtrlSetState($hNoPDFs, $GUI_UNCHECKED)
						Else
							$aHandler = StringSplit($sHandler, "\")
							GUICtrlSetData($hPDFPath, $aHandler[$aHandler[0]])
						EndIf
					Else
						GUICtrlSetData($hPDFPath, "")
					EndIf

				Case $hMsg = $hInstall
					If $bUpdate Then RunRemoval(True)

					$aConfig[$vMode] = _IsChecked($hActive)

					$aSettings[$bNoApps] = _IsChecked($hNoApps)
					$aSettings[$bNoBing] = _IsChecked($hSearch)
					$aSettings[$bNoPDFs] = _IsChecked($hNoPDFs)
					$aSettings[$bNoTray] = _IsChecked($hNoIcon)
					$aSettings[$sPDFApp] = $sHandler
					$aSettings[$sSearch] = GUICtrlRead($hEngine)
					$aSettings[$sSearchPath] = $sEngine
					$aSettings[$bStartup] = _IsChecked($hStartup)

					GUISetState(@SW_HIDE, $hSettings)
					RunInstall($aConfig, $aSettings)
					SetAppRegistry($aConfig[$vMode])
					If $aConfig[$vMode] Then
						For $iLoop = 0 To 3 Step 1
							$aChannels[$iLoop] = _IsChecked($hChannels[$iLoop])
						Next
						SetIFEORegistry($aChannels)
					EndIf
					GUISetState(@SW_SHOW, $hFinish)

				Case Else
					;;;

			EndSelect

		WEnd

	EndIf

EndFunc

Func SetAppRegistry($bAllUsers)

	Local $sHive = ""
	Local $sLocation = ""

	If $bAllUsers Then
		$sLocation = "C:\Program Files\MSEdgeRedirect\"
		If _WinAPI_IsWow64Process() Then
			$sHive = "HKLM64"
		Else
			$sHive = "HKLM"
		EndIf
	Else
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
		If _WinAPI_IsWow64Process() Then
			$sHive = "HKCU64"
		Else
			$sHive = "HKCU"
		EndIf
	EndIf

	; App Paths
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "", "REG_SZ", $sLocation & "MSEdgeRedirect.exe")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "Path", "REG_SZ", $sLocation)

	; URI Handler for Pre Win11 22494 Installs
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "SupportedProtocols", "REG_SZ", "microsoft-edge")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge", "", "REG_SZ", "URL:microsoft-edge")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge", "URL Protocol", "REG_SZ", "")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge\shell\open\command", "", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe" "%1"')

	; Generic Program Info
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect\DefaultIcon", "", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe",0')
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect\shell\open\command", "", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe" "%1"')
	RegWrite($sHive & "\Software\Classes\Applications\MSEdgeRedirect.exe", "FriendlyAppName", "REG_SZ", "MSEdgeRedirect")
	RegWrite($sHive & "\Software\Classes\Applications\MSEdgeRedirect.exe\DefaultIcon", "", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe",0')

	; Uninstall Info
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayIcon", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe",0')
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayName", "REG_SZ", "MSEdgeRedirect")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayVersion", "REG_SZ", $sVersion)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "EstimatedSize", "REG_DWORD", 1536)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "HelpLink", "REG_SZ", "https://msedgeredirect.com/wiki")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "InstallDate", "REG_SZ", StringReplace(_NowCalcDate(), "/", ""))
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "InstallLocation", "REG_SZ", $sLocation)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Language", "REG_DWORD", 1033)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "ModifyPath", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe" /change')
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "NoModify", "REG_DWORD", 0)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "NoRepair", "REG_DWORD", 1)
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Publisher", "REG_SZ", "Robert Maehl Software")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Readme", "REG_SZ", "https://msedgeredirect.com/blob/main/README.md")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "UninstallString", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe" /uninstall')
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "URLInfoAbout", "REG_SZ", "https://msedgeredirect.com")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "URLUpdateInfo", "REG_SZ", "https://msedgeredirect.com/releases")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Version", "REG_SZ", $sVersion)

EndFunc

Func SetIFEORegistry(ByRef $aChannels)

	Local $sHive = ""

	If _WinAPI_IsWow64Process() Then
		$sHive = "HKLM64"
	Else
		$sHive = "HKLM"
	EndIf

	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe", "UseFilter", "REG_DWORD", 1)
	For $iLoop = 1 To $aEdges[0] Step 1
		If $aChannels[$iLoop - 1] Then
			RegWrite($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop)
			RegWrite($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop, "Debugger", "REG_SZ", "C:\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe")
			RegWrite($sHive & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop, "FilterFullPath", "REG_SZ", $aEdges[$iLoop])
			FileCopy($aEdges[$iLoop], StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe"), $FC_OVERWRITE)
		EndIf
	Next
EndFunc

Func SetOptionsRegistry($sName, $vValue, $bAllUsers, $bManaged = False)

	Local Static $sHive = ""
	Local Static $sPolicy = ""
	#forceref $sHive

	If $sHive = "" Then
		If $bAllUsers Then
			If _WinAPI_IsWow64Process() Then
				$sHive = "HKLM64"
			Else
				$sHive = "HKLM"
			EndIf
		Else
			If _WinAPI_IsWow64Process() Then
				$sHive = "HKCU64"
			Else
				$sHive = "HKCU"
			EndIf
		EndIf

		If $bManaged Then $sPolicy = "Policies\"
	EndIf

	Select
		Case IsBool($vValue)
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_DWORD", $vValue)

		Case IsString($vValue)
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_SZ", $vValue)

		Case Else
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_SZ", $vValue)
	EndSelect

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

Func _Bool($sString)
	If $sString = "True" Then
		Return True
	ElseIf $sString = "False" Then
		Return False
	Else
		Return $sString
	EndIf
EndFunc

Func _ChangeSearchEngine($sURL)

	If StringInStr($sURL, "bing.com/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(.*)(q=)", "")

		Switch _GetSettingValue("Search")

			Case "Ask"
				Return "https://www.ask.com/web?q=" & $sURL

			Case "Baidu"
				Return "https://www.baidu.com/s?wd=" & $sURL

			Case "Custom"
				Return _GetSettingValue("SearchPath") & $sURL

			Case "DuckDuckGo"
				Return "https://duckduckgo.com/?q=" & $sURL

			Case "Ecosia"
				Return "https://www.ecosia.org/search?q=" & $sURL

			Case "Google"
				Return "https://www.google.com/search?q=" & $sURL

			Case "Sogou"
				Return "https://www.sogou.com/web?query=" & $sURL

			Case "Yahoo"
				Return "https://search.yahoo.com/search?p=" & $sURL

			Case "Yandex"
				Return "https://yandex.com/search/?text=" & $sURL

			Case Null
				Return "https://bing.com/search?q=" & $sURL

			Case Else
				Return _GetSettingValue("SearchPath") & $sURL

		EndSwitch
	Else
		Return $sURL
	EndIf

EndFunc


Func _DecodeAndRun($sCMDLine)

	Local $sCaller
	Local $aLaunchContext

	Select
		Case StringInStr($sCMDLine, "--default-search-provider=?")
			FileWrite($hLogs[2], _NowCalc() & " - Skipped Settings URL: " & $sCMDLine & @CRLF)
		Case StringInStr($sCMDLine, ".pdf") And _GetSettingValue("NoPDFs")
			ShellExecute(_GetSettingValue("PDFApp"), $sCMDLine)
		Case StringInStr($sCMDLine, "--app-id") And _GetSettingValue("NoApps") ; TikTok and other Apps
			$sCMDLine = StringRegExpReplace($sCMDLine, "(.*)(--app-fallback-url=)", "")
			$sCMDLine = StringRegExpReplace($sCMDLine, "(?= --)(.*)", "")
			ShellExecute($sCMDLine)
		Case StringInStr($sCMDLine, "Windows.Widgets")
			$sCaller = "Windows.Widgets"
			ContinueCase
		Case StringRegExp($sCMDLine, "microsoft-edge:[\/]*?\?launchContext1")
			$aLaunchContext = StringSplit($sCMDLine, "=")
			If $aLaunchContext[0] >= 3 Then
				If $sCaller = "" Then $sCaller = $aLaunchContext[2]
				FileWrite($hLogs[1], _NowCalc() & " - Redirected Edge Call from: " & $sCaller & @CRLF)
				$sCMDLine = _UnicodeURLDecode($aLaunchContext[$aLaunchContext[0]])
				If _WinAPI_UrlIs($sCMDLine) Then
					If _GetSettingValue("NoBing") Then $sCMDLine = _ChangeSearchEngine($sCMDLine)
					ShellExecute($sCMDLine)
				Else
					FileWrite($hLogs[2], _NowCalc() & " - Invalid Regexed URL: " & $sCMDLine & @CRLF)
				EndIf
			Else
				FileWrite($hLogs[2], _NowCalc() & " - Command Line Missing Needed Parameters: " & $sCMDLine & @CRLF)
			EndIf
		Case Else
			$sCMDLine = StringRegExpReplace($sCMDLine, "(.*) microsoft-edge:[\/]*", "")
			If _WinAPI_UrlIs($sCMDLine) Then
				If _GetSettingValue("NoBing") Then $sCMDLine = _ChangeSearchEngine($sCMDLine)
				ShellExecute($sCMDLine)
			Else
				FileWrite($hLogs[2], _NowCalc() & " - Invalid URL: " & $sCMDLine & @CRLF)
			EndIf
	EndSelect
EndFunc


Func _GetSettingValue($sSetting, $bPortable = False)

	Local $vReturn = Null

	Local $sHive1 = ""
	Local $sHive2 = ""

	If _WinAPI_IsWow64Process() Then
		$sHive1 = "HKLM64"
		$sHive2 = "HKCU64"
	Else
		$sHive1 = "HKLM"
		$sHive2 = "HKCU"
	EndIf

	Select

		Case RegRead($sHive1 & "\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead($sHive1 & "\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD
					$vReturn =  Number(RegRead($sHive1 & "\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[0], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD
					$vReturn = Number(RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[0], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD
					$vReturn = Number(RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[0], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case Not IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, Null) = Null And Not $bPortable
			$vReturn = _Bool(IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, False))

		Case Not IniRead(@ScriptDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, Null) = Null
			$vReturn = _Bool(IniRead(@ScriptDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, False))

		Case Else
			;;;

	EndSelect

	Return $vReturn

EndFunc

Func _IsInstalled()

	Local $sHive1 = ""
	Local $sHive2 = ""
	Local $aReturn[3] = [False, "", ""]

	If _WinAPI_IsWow64Process() Then
		$sHive1 = "HKLM64"
		$sHive2 = "HKCU64"
	Else
		$sHive1 = "HKLM"
		$sHive2 = "HKCU"
	EndIf

	$aReturn[2] = RegRead($sHive1 & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayVersion")
	If @error Then
		$aReturn[2] = RegRead($sHive2 & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayVersion")
		If @error Then
			;;;
		Else
			$aReturn[0] = True
			$aReturn[1] = $sHive2
		EndIf
	Else
		$aReturn[0] = True
		$aReturn[1] = $sHive1
	EndIf

	Return $aReturn

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

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

; #FUNCTION# ====================================================================================================================
; Name ..........: _UnicodeURLDecode
; Description ...: Tranlates a URL-friendly string to a normal string
; Syntax ........: _UnicodeURLDecode($toDecode)
; Parameters ....: $$toDecode           - The URL-friendly string to decode
; Return values .: The URL decoded string
; Author ........: nfwu, Dhilip89
; Modified ......: 11/17/2021
; Remarks .......: Modified from _URLDecode() that only supported non-unicode.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
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
