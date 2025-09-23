#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\MSEdgeRedirect.ico
#AutoIt3Wrapper_Outfile=MSEdgeRedirect_x86.exe
#AutoIt3Wrapper_Outfile_x64=MSEdgeRedirect.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=https://www.msedgeredirect.com
#AutoIt3Wrapper_Res_CompanyName=Robert Maehl Software
#AutoIt3Wrapper_Res_Description=MSEdgeRedirect
#AutoIt3Wrapper_Res_Fileversion=0.8.0.0
#AutoIt3Wrapper_Res_ProductName=MSEdgeRedirect App & Service
#AutoIt3Wrapper_Res_ProductVersion=0.8.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Compatibility=Win10
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 -v1 -v2 -v3
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so
#AutoIt3Wrapper_Res_Icon_Add=Assets\MSEdgeRedirect.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <Misc.au3>
#include <Array.au3>
#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>
#include <EditConstants.au3>
#include <TrayConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>

#include "Includes\_Compat.au3"
#include "Includes\_Logging.au3"
#include "Includes\_Theming.au3"
#include "Includes\_Security.au3"
#include "Includes\_Settings.au3"
#include "Includes\_Translation.au3"
#include "Includes\_URLModifications.au3"

#include "Includes\Base64.au3"
#include "Includes\ResourcesEx.au3"

Opt("TrayMenuMode", 3)
Opt("TrayAutoPause", 0)
Opt("GUICloseOnESC", 0)

#include "MSEdgeRedirect_Wrapper.au3"

SetupAppdata()
ProcessCMDLine()

Func ActiveMode(ByRef $aCMDLine)

	Local $iIndex
	Local $sCMDLine = ""
	Local $sParent = _WinAPI_GetProcessName(_WinAPI_GetParentProcess())

	$aCMDLine = FixTreeIntegrity($aCMDLine)
	CheckEdgeIntegrity($aCMDLine[1])
	$aCMDLine[1] = StringReplace($aCMDLine[1], "msedge.exe", "msedge_IFEO.exe")
	
	Select
		Case $aCMDLine[0] = 1 ; No Parameters
			ContinueCase
		Case $aCMDLine[0] = 2 And UBound($aCMDLine) < 2
			ReDim $aCMDLine[3]
			$aCMDLine[2] = ""
			ContinueCase
		Case $aCMDLine[0] = 2 And FileExists($aCMDLine[2])
			If FileExists($aCMDLine[2]) Then $aCMDLine[2] = '"' & $aCMDLine[2] & '"'
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--uninstall" ; Uninstalling Edge
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--suspend-background-mode" ; Uninstalling Edge
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--from-installer" ; Installing Edge
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "-inprivate" ; In Private Browsing, Short Flag, No Parameters
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--inprivate" ; In Private Browsing, No Parameters
			ContinueCase
		Case _ArraySearch($aCMDLine, "--winrt-background-task-event", 2, 0, 0, 1) > 0 ; #94 & #95, Apps
			ContinueCase
		Case _ArraySearch($aCMDLine, "--web-widget-jumplist-launch", 2, 0, 0, 1) > 0 ; #123, EdgeBar
			ContinueCase
		Case _ArraySearch($aCMDLine, "--notification-launch-id", 2, 0, 0, 1) > 0 ; #225, Web App Notifications
			ContinueCase
		Case _ArraySearch($aCMDLine, "--app-id", 2, 0, 0, 1) > 0 And Not _GetSettingValue("NoApps", "Bool")
			ContinueCase
		Case _ArraySearch($aCMDLine, "--remote-debugging-port=", 2, 0, 0, 1) > 0 ; #271, Debugging Apps
			ContinueCase
		Case _ArraySearch($aCMDLine, "--profile-directory=", 2, 0, 0, 1) > 0 ; #68, Multiple Profiles
			ContinueCase
		Case _ArraySearch($aCMDLine, "--user-data-dir=", 2, 0, 0, 1) > 0 ; #463, Multiple Profiles
			ContinueCase			
		Case $sParent = "MSEdgeRedirect.exe"
			$iIndex = _ArraySearch($aCMDLine, "--from-ie-to-edge", 2, 0, 0, 1)
			If $iIndex Then
				_ArrayDelete($aCMDLine, $iIndex)
				$sCMDLine = _ArrayToString($aCMDLine, " ", 2, -1)
				_DecodeAndRun(Default, $sCMDLine)
			EndIf
			$sCMDLine = _ArrayToString($aCMDLine, " ", 2, -1)
			_SafeRun($aCMDLine[1], $sCMDLine)
		Case _DoesParentProcessWantEdge($sParent)
			ContinueCase
		Case $aCMDLine[0] = 2 And $aCMDLine[2] = "--continue-active-setup"
			_SafeRun($aCMDLine[1], $aCMDLine[2])
		Case _IsURLLocalHost($aCMDLine)
			$sCMDLine = _ArrayToString($aCMDLine, " ", 2, -1)
			_Log($hLogs[$URIFailures], "Skipped Localhost URL: " & $sCMDLine & @CRLF)
		Case Else
			$sCMDLine = _ArrayToString($aCMDLine, " ", 2, -1)
			_DecodeAndRun($aCMDLine[1], $sCMDLine)
	EndSelect

EndFunc

Func CheckEdgeIntegrity($sLocation)
	If StringInStr($sLocation, "ie_to_edge_stub") Then
		;;;
	ElseIf $sLocation = "" Then
		_LogClose()
		Exit
	Else
		Select
			Case Not FileExists(StringReplace($sLocation, "\msedge.exe", "\msedge_IFEO.exe"))
				If WinExists(_Translate($aMUI[1], "Admin File Copy Required")) Then
					_LogClose()
					Exit ; #202
				EndIf
				If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _
					_Translate($aMUI[1], "Admin Rights Required"), _
					_Translate($aMUI[1], "The IFEO junctions for MSEdgeRedirect are missing and need to be created. Create Now?"), _
					0) = $IDYES Then ShellExecuteWait(@ScriptFullPath, "/repair", @ScriptDir, "RunAs")
				If @error Then MsgBox($MB_ICONERROR+$MB_OK, _
					_Translate($aMUI[1], "Copy Failed"), _
					_Translate($aMUI[1], "Unable to create the IFEO junction without Admin Rights!"))
			Case Else
				;;;
		EndSelect
	EndIf
EndFunc

Func FixTreeIntegrity($aCMDLine)

	Local $iParent = _WinAPI_GetParentProcess()

	Switch _WinAPI_GetProcessName($iParent)
		
		Case "MSEdge.exe"

			_Log($hLogs[$AppGeneral], "" & "Caught MSEdge Parent Process, Launched by " & _WinAPI_GetProcessName(_WinAPI_GetParentProcess($iParent)) & ", Grabbing Parameters." & @CRLF)

			Local $aAdjust

			; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
			Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))

			_WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

			Redim $aCMDLine[2]
			$aCMDLine[0] = 0
			$aCMDLine[1] = _WinAPI_GetProcessFileName($iParent)

			_ArrayConcatenate($aCMDLine, StringSplit(_WinAPI_GetProcessCommandLine($iParent), " ", $STR_NOCOUNT))

			$aCMDLine[0] = UBound($aCMDLine) - 1

			ProcessClose($iParent)
			_WinAPI_CloseHandle($hToken)

		Case "MSEdgeRedirect.exe"

			;;;

		Case Else

			;;;

	EndSwitch
	Return $aCMDLine

EndFunc

Func ProcessCMDLine()

	Local $aPIDs
	Local $bHide = _GetSettingValue("NoTray", "Bool")
	Local $hFile = @ScriptDir & ".\Setup.ini"
	Local $bForce = False
	; Local $iChance = 10
	Local $iParams = $CmdLine[0]
	Local $sCMDLine = _ArrayToString($CmdLine, " ", 1)
	Local $bSilent = False
	Local $aInstall[3] ; [Installed, Registry Hive, Version]

	$aInstall = _IsInstalled()
	If DriveGetType(@ScriptDir) = "Removable" And FileExists(".\Settings.ini") Then _GetSettingValue("SetPortable")

	If $iParams > 0 Then

		$CMDLine = RepairCMDLine($CMDLine)
		If _ArraySearch($aEdges, $CMDLine[1]) > 0 Or StringInStr($CMDLine[1], "ie_to_edge_stub.exe") Then ; Image File Execution Options Mode
			ActiveMode($CMDLine)
			; TODO: Parse $aSettings[], decrease likelyhood based on number of enabled features so that users with more features enabled aren't spammed
			; TODO: Revamp $aSettings to remove "Custom", have <whatever>PATH to replace "CUSTOM"
			If Not _GetSettingValue("NoUpdates", "Bool") And Random(1, 10, 1) = 1 Then RunUpdateCheck()
			_LogClose()
			Exit
		ElseIf StringLeft($CMDLine[1], 5) = "MSER:" Then ; Future Edge Add-On Mode
			;;;
		EndIf

		Do
			Switch $CmdLine[1]
				Case "/?", "/help"
					MsgBox(0, "Help and Flags", _
							"MSEdgeRedirect" & @CRLF & _
							@CRLF & _
							@TAB & "/admin    " & @TAB & "Attempts to run MSEdgeRedirect as admin" & @CRLF & _
							@TAB & "/change   " & @TAB & "Reruns Installer" & @CRLF & _
							@TAB & "/hide     " & @TAB & "Hides the tray icon" & @CRLF & _
							@TAB & "/force    " & @TAB & "Skips Safety Checks" & @CRLF & _
							@TAB & "/kill     " & @TAB & "Kills other MSEdgeRedirect processes" & @CRLF & _
							@TAB & "/portable " & @TAB & "Runs MSEdgeRedirect in portable mode" & @CRLF & _
							@TAB & "/repair   " & @TAB & "Repairs IFEO directory junctions" & @CRLF & _
							@TAB & "/settings " & @TAB & "Opens Settings Menu" & @CRLF & _
							@TAB & "/si       " & @TAB & "Runs a Silent Install" & @CRLF & _
							@TAB & "/update   " & @TAB & "Downloads the latest RELEASE (default) or DEV build" & @CRLF & _
							@TAB & "/uninstall" & @TAB & "Uninstalls MSEdgeRedirect" & @CRLF & _
							@CRLF & _
							@CRLF)
					_LogClose()
					Exit 0
				Case "/admin"
					If Not $bIsAdmin Then
						ShellExecute(@ScriptFullPath, $sCMDLine, @ScriptDir, "RunAs")
						_LogClose()
						Exit
					Else
						_ArrayDelete($CmdLine, 1)
					EndIf
				Case "/change"
					RunSetup(True, $bSilent, 1)
					_LogClose()
					Exit
				Case "/ContinueActive"
					If Not $bIsAdmin Then
						MsgBox($MB_ICONERROR+$MB_OK, _
							_Translate($aMUI[1], "Admin Required"), _
							_Translate($aMUI[1], "Unable to install Active Mode without Admin Rights!"))
						_Log($hLogs[$AppFailures], "" & "Active Mode UAC Elevation Attempt Failed!" & @CRLF)
						_LogClose()
						Exit
					Else
						RunSetup($aInstall[0], False, -2)
					EndIf
				Case "/ContinueEurope", "/SetEurope"
					Select
						Case Not $bIsAdmin
							MsgBox($MB_ICONERROR+$MB_OK, _
								_Translate($aMUI[1], "Admin Required"), _
								_Translate($aMUI[1], "Unable to Setup Europe Mode without Admin Rights!"))
							_Log($hLogs[$AppFailures], "" & "Europe Mode UAC Elevation Attempt Failed!" & @CRLF)
							_LogClose()
							Exit
						Case Not RegRead("HKLM\SYSTEM\CurrentControlSet\Services\UCPD", "Start") = 4
							ContinueCase
						Case Not RegRead("HKLM\SYSTEM\CurrentControlSet\Services\UCPD", "FeatureV2") = 0
							If MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, _
								_Translate($aMUI[1], "Reboot Required"), _
								_Translate($aMUI[1], "A Reboot/Restart is required to disable User Choice Protection Driver (UCPD), would you like to do so now?")) = $IDYES Then
								RegWrite("HKLM\SYSTEM\CurrentControlSet\Services\UCPD", "Start", "REG_DWORD", 4)
								RegWrite("HKLM\SYSTEM\CurrentControlSet\Services\UCPD", "FeatureV2", "REG_DWORD", 0)
								RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe")
								RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe", "UseFilter", "REG_DWORD", 1)
								RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe\MSER")
								RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe\MSER", "Debugger", "REG_SZ", $sDrive & "\Windows\System32\ping.exe")
								RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\UCPDMgr.exe\MSER", "FilterFullPath", "REG_SZ", $sDrive & "\Windows\System32\UCPDMgr.exe")								
								Shutdown($SD_REBOOT)
							EndIf
							_LogClose()
							Exit
						Case Else
							RunSetup($aInstall[0], False, -5)
					EndSelect
				Case "/f", "/force"
					$bForce = True
					_ArrayDelete($CmdLine, 1)
				Case "/h", "/hide"
					$bHide = True
					_ArrayDelete($CmdLine, 1)
				Case "/kill"
					$aPIDs = ProcessList(@ScriptName)
					For $iLoop = 1 To $aPIDs[0][0] Step 1
						If $aPIDs[$iLoop][1] <> @AutoItPID Then ProcessClose($aPIDs[$iLoop][1])
					Next
					_LogClose()
					Exit
				Case "/p", "/portable"
					_GetSettingValue("SetPortable")
					_ArrayDelete($CmdLine, 1)
				Case "/repair"
					RunRepair()
					_LogClose()
					Exit
				Case "/settings"
					$aPIDs = ProcessList(@ScriptName)
					For $iLoop = 1 To $aPIDs[0][0] Step 1
						If $aPIDs[$iLoop][1] <> @AutoItPID Then ProcessClose($aPIDs[$iLoop][1])
					Next
					RunSetup(2, False, 2)
					If Not $bIsPriv Then ShellExecute(@ScriptFullPath)
					_LogClose()
					Exit
				Case "/si", "/silentinstall"
					$bSilent = True
					Select
						Case UBound($CmdLine) = 2
							_ArrayDelete($CmdLine, 1)
						Case UBound($CmdLine) > 2 And FileExists($CmdLine[2])
							$hFile = $CmdLine[2]
							_ArrayDelete($CmdLine, "1-2")
						Case StringLeft($CmdLine[2], 1) = "/"
							_ArrayDelete($CmdLine, 1)
						Case Else
							MsgBox(0, _
								"Invalid", _
								'Invalid file - "' & $CmdLine[2] & @CRLF)
							_LogClose()
							Exit 87 ; ERROR_INVALID_PARAMETER
					EndSelect
				#cs
				Case "/u", "/update"
					Select
						Case UBound($CmdLine) = 2
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, 1)
						Case UBound($CmdLine) > 2 And $CmdLine[2] = "dev"
							InetGet("https://nightly.link/rcmaehl/MSEdgeRedirect/workflows/MSER/main/mser.zip", @ScriptDir & "\MSEdgeRedirect_dev.zip")
							_ArrayDelete($CmdLine, "1-2")
						Case UBound($CmdLine) > 2 And $CmdLine[2] = "release"
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, "1-2")
						Case StringLeft($CmdLine[2], 1) = "/"
							InetGet("https://fcofix.org/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe", @ScriptDir & "\MSEdgeRedirect_Latest.exe")
							_ArrayDelete($CmdLine, 1)
						Case Else
							MsgBox(0, _
								"Invalid", _
								'Invalid release type - "' & $CmdLine[2] & "." & @CRLF)
							Exit 87 ; ERROR_INVALID_PARAMETER
					EndSelect
				#ce
				Case "/uninstall"
					RunRemoval()
					_LogClose()
					Exit
				Case "/wingetinstall"
					If Not $bIsAdmin Then
						ShellExecute(@ScriptFullPath, $sCMDLine, @ScriptDir, "RunAs")
						If @error Then
							;;;
						Else
							_LogClose()
							Exit
						EndIf
					EndIf
					$bSilent = True
					$hFile = "WINGET"
					_ArrayDelete($CmdLine, 1)
				Case "--MSEdgeRedirect"
					_Log($hLogs[$PEBIAT], "" & "Recursion Prevention Check Failed. " & @CRLF & _
						"Commandline: " & _ArrayToString($CmdLine) & @CRLF & _ 
						"Parent: " & _WinAPI_GetProcessName(_WinAPI_GetParentProcess()) & @CRLF)
					_ArrayDelete($CmdLine, 1)
				Case Else
					_Log($hLogs[$PEBIAT], "" & "Unexpected Commandline: " & _ArrayToString($CmdLine) & @CRLF)
					If @Compiled Then ; support for running non-compiled script - mLipok
						MsgBox(0, _
							_Translate($aMUI[1], "Invalid"), _
							_Translate($aMUI[1], 'Invalid parameter - "') & $CmdLine[1] & '".' & @CRLF)
						_LogClose()
						Exit 87 ; ERROR_INVALID_PARAMETER
					EndIf
			EndSwitch
		Until UBound($CmdLine) <= 1
	Else
		;;;
	EndIf

	If $hFile = "WINGET" Then
		;;;
	ElseIf Not $bForce Then
		RunArchCheck($bSilent)
	Else
		;;;
	EndIf

	If Not _GetSettingValue("IsPortable") Then

		Select
			Case Not $aInstall[0] ; Not Installed
				RunSetup(False, $bSilent, 0, $hFile)
			Case _VersionCompare($sVersion, $aInstall[2]) ; Installed, Out of Date
				Select
					Case StringInStr($aInstall[1], "HKCU") ; Installed, Service Mode
						RunSetup($aInstall[0], $bSilent, 0, $hFile)
					Case StringInStr($aInstall[1], "HKLM") And Not $bIsAdmin And @Compiled; Installed, Active Mode, Not Admin
						ShellExecute(@ScriptFullPath, $sCMDLine, @ScriptDir, "RunAs")
						If @error Then
							If Not $bSilent Then MsgBox($MB_ICONWARNING+$MB_OK, _
								_Translate($aMUI[1], "Existing Active Mode Install"), _
								_Translate($aMUI[1], "Unable to update an existing Active Mode install without Admin Rights! The installer will continue however."))
							ContinueCase
						Else
							_LogClose()
							Exit
						EndIf
					Case StringInStr($aInstall[1], "HKLM") ; Installed, Active Mode
						RunSetup($aInstall[0], $bSilent, 0, $hFile)
				EndSelect
			Case StringInStr($aInstall[1], "HKCU") ; Installed, Up to Date, Service Mode
				If @ScriptDir <> @LocalAppDataDir & "\MSEdgeRedirect" Then
					RunSetup($aInstall[0], $bSilent, 0, $hFile)
					;ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", "", @LocalAppDataDir & "\MSEdgeRedirect\")
				Else
					$aPIDs = ProcessList(@ScriptName)
					For $iLoop = 1 To $aPIDs[0][0] Step 1
						If $aPIDs[$iLoop][1] <> @AutoItPID Then
							$bHide = False
							ProcessClose($aPIDs[$iLoop][1])
						EndIf
					Next
				EndIf
			Case StringInStr($aInstall[1], "HKLM") And Not $bIsAdmin ; Installed, Up to Date, Active Mode, Not Admin
				ShellExecute(@ScriptFullPath, $sCMDLine, @ScriptDir, "RunAs")
				If @error Then
					If Not $bSilent Then MsgBox($MB_ICONWARNING+$MB_OK, _
						_Translate($aMUI[1], "Existing Active Mode Install"), _
						_Translate($aMUI[1], "Unable to update an existing Active Mode install without Admin Rights! The installer will continue however."))
					ContinueCase
				Else
					_LogClose()
					Exit
				EndIf
			Case Else
				RunSetup(True, $bSilent, 0, $hFile)
		EndSelect
	EndIf
	RunHTTPCheck()
	ReactiveMode($bHide)

EndFunc

Func ReactiveMode($bHide = False)

	Local $aAdjust

	Local $hMsg

	Global $oWMI  = ObjGet("winmgmts:\\")
	Global $oSink = ObjCreate("WbemScripting.SWbemSink")
	_registerProcessCreation()

	; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
	Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))

	_WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

	TrayCreateItem($sVersion)
	TrayItemSetState(-1, $TRAY_DISABLE)
	TrayCreateItem("")
	Local $hSettings = TrayCreateItem("Settings")
	Local $hDonate = TrayCreateItem("Donate")
	TrayCreateItem("")
	Local $hUpdate = TrayCreateItem("Check for Updates")
	Local $hExit = TrayCreateItem("Exit")

	If $bHide Then TraySetState($TRAY_ICONSTATE_HIDE)

	Local $sRegex
	Local $iSIHost = ProcessExists("sihost.exe")
	Local $bHavePath = True
	Local $aProcessList
	Local $sProcessPath
	Local $sCommandline	

	If _GetSettingValue("NoApps", "Bool") Then
		$sRegex = "(?i).*(microsoft\-edge|app\-id).*"
	Else
		$sRegex = "(?i).*(microsoft\-edge).*"
	EndIf

	While True
		$hMsg = TrayGetMsg()

#cs
		$aProcessList = _WinAPI_EnumChildProcess($iSIHost)
		If Not @error Then
			ProcessClose($aProcessList[1][0])
			$sCommandline = _WinAPI_GetProcessCommandLine($aProcessList[1][0])
			$sProcessPath = _WinAPI_GetProcessFileName($aProcessList[1][0])
			If @error Then $bHavePath = False
			If StringRegExp($sCommandline, $sRegex) Then _DecodeAndRun(Default, $sCommandline)
			;Relaunch other processes without SIHOST Parent
			If Not StringRegExp($sCommandline, $sRegex) And $bHavePath = True Then _SafeRun($sProcessPath, $sCommandline)
		Else
			$aProcessList = ProcessList("msedge.exe")
			For $iLoop = 1 To $aProcessList[0][0] - 1
				$sCommandline = _WinAPI_GetProcessCommandLine($aProcessList[$iLoop][1])
				If Not StringRegExp($sCommandline, $sRegex) Then ContinueLoop
				ProcessClose($aProcessList[$iLoop][1])
				_DecodeAndRun(Default, $sCommandline)
			Next
		EndIf
#ce

		Switch $hMsg

			Case $hSettings
				ShellExecute(@ScriptFullPath, "/settings", @ScriptDir)

			Case $hExit
				ExitLoop

			Case $hDonate
				ShellExecute("https://www.paypal.com/donate/?hosted_button_id=YL5HFNEJAAMTL")

			Case $hUpdate
				RunUpdateCheck(True)

			Case Else

		EndSwitch
	WEnd

	_WinAPI_AdjustTokenPrivileges($hToken, $aAdjust, 0, $aAdjust)
	_WinAPI_CloseHandle($hToken)
	_LogClose()
	Exit

EndFunc

Func SINK_OnObjectReady($oProcess)
    ProcessClose($oProcess.TargetInstance.ProcessID)
	Local $sRegex = "(?i).*(microsoft\-edge|app\-id).*"
	$sCommandline = _WinAPI_GetProcessCommandLine($oProcess.TargetInstance.ProcessID)
	$sProcessPath = _WinAPI_GetProcessFileName($oProcess.TargetInstance.ProcessID)
	ConsoleWrite($sCommandline & @CRLF)
	If StringRegExp($sCommandline, $sRegex) Then _DecodeAndRun(Default, $sCommandline)
EndFunc

Func _registerProcessCreation()
    ; Events with the prefix “SINK_” are linked to corresponding AutoIt functions (we only need SINK_OnObjectReady).
    ObjEvent($oSink, "SINK_")

    ; Queries the __InstanceCreationEvent events on the WMI class Win32_Process every 100 ms.
    $oWMI.ExecNotificationQueryAsync($oSink, "SELECT * FROM __InstanceCreationEvent WITHIN 0.1 WHERE TargetInstance ISA 'Win32_Process' AND (TargetInstance.name = 'msedge.exe')")
EndFunc

Func RepairCMDLine($aCMDLine)

	Local $sCMDLine
	Local $sDelim = _ArraySafeDelim($aCMDLine)

	$sCMDLine = _ArrayToString($aCMDLine, $sDelim)
	Select
		Case StringInStr($sCMDLine, "Program" & $sDelim & "Files" & $sDelim & "(x86)")
			$sCMDLine = StringReplace($sCMDLine, "Program" & $sDelim & "Files" & $sDelim & "(x86)", "Program Files (x86)")
		Case StringInStr($sCMDLine, $sDelim & "--" & $sDelim)
			$sCMDLine = StringReplace($sCMDLine, "--" & $sDelim, "")
		Case Else
		;;;
	EndSelect

	$aCMDLine = StringSplit($sCMDLine, $sDelim, $STR_ENTIRESPLIT+$STR_NOCOUNT)
	$aCMDLine[0] = UBound($aCMDLine) - 1

	Return $aCMDLine

EndFunc

Func RunArchCheck($bSilent = False)
	If @Compiled And $bIsWOW64 Then
		If Not $bSilent Then
			MsgBox($MB_ICONERROR+$MB_OK, _
				_Translate($aMUI[1], "Wrong Version"), _
				_Translate($aMUI[1], "The 64-bit Version of MSEdgeRedirect must be used with 64-bit Windows!"))
		EndIf
		_Log($hLogs[$AppFailures], "" & "32 Bit Version on 64 Bit System. EXITING!" & @CRLF)
		_LogClose()
		Exit 216 ; ERROR_EXE_MACHINE_TYPE_MISMATCH
	EndIf
EndFunc

Func RunHTTPCheck($bSilent = False)

	Local $aDefaults[3]
	Local Enum $hHTTP, $hHTTPS, $hMSEdge

	$aDefaults[$hHTTP] = RegRead("HKCU\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", "ProgId")
	If @error Then Return
	$aDefaults[$hHTTPS] = RegRead("HKCU\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice", "ProgId")
	If @error Then Return
	$aDefaults[$hMSEdge] = RegRead("HKCU\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\microsoft-edge\UserChoice", "ProgId")
	If @error Then Return

	If StringInStr($aDefaults[$hMSEdge], "MSEdge") Then
		If $aDefaults[$hHTTP] = $aDefaults[$hMSEdge] Or $aDefaults[$hHTTPS] = $aDefaults[$hMSEdge] Then
			If Not $bSilent Then
				MsgBox($MB_ICONERROR+$MB_OK, _
					_Translate($aMUI[1], "Edge Set As Default"), _
					_Translate($aMUI[1], "You must set a different Default Browser to use MSEdgeRedirect! Once this is corrected, please relaunch MSEdgeRedirect."))
			EndIf
			_Log($hLogs[$AppFailures], "" & "Found same MS Edge for both default browser and microsoft-edge handling, EXITING!" & @CRLF)
			_LogClose()
			Exit 4315 ; ERROR_MEDIA_INCOMPATIBLE
		EndIf
	EndIf

EndFunc

Func _DecodeAndRun($sEdge = $aEdges[1], $sCMDLine = "")

	Local $sURL = ""
	Local $aCMDLine

	Select

		; Run Edge
		Case StringRegExp($sCMDLine, "--default-search-provider=\? --out-pipe-name=MSEdgeDefault[a-z0-9]+")
			_Log($hLogs[$AppSecurity], "Passed Through MS-Settings Call: " & $sCMDLine & @CRLF)
			_SafeRun($sEdge, $sCMDLine)
		Case $sCMDLine = "--no-startup-window --win-session-start"
			_Log($hLogs[$AppSecurity], "Passed Through MSEdge Startup Call: " & $sCMDLine & @CRLF)
			_SafeRun($sEdge, $sCMDLine)

		; Run Another App
		Case FileExists(StringReplace($sCMDLine, "--single-argument ", "")); File Handling
			If _GetSettingValue("NoFiles", "Bool") Or _GetSettingValue("NoPDFs", "Bool") Then
				$sCMDLine = StringReplace($sCMDLine, "--single-argument ", "")
				If _IsSafeFile($sCMDLine) Then ShellExecute('"' & $sCMDLine & '"', "", "", $SHEX_EDIT)
			EndIf
		#cs
			If _GetSettingValue("NoPDFs", "Bool") Then
				$sCMDLine = StringReplace($sCMDLine, "--single-argument ", "")
				Switch _GetSettingValue("PDFApp", "String")
					Case "Default", False
						If RunPDFCheck() And _IsSafePDF($sCMDLine) Then ShellExecute('"' & $sCMDLine & '"')
					Case Else
						ShellExecute(_GetSettingValue("PDFApp", "String"), '"' & $sCMDLine & '"')
				EndSwitch
			Else
				$sCMDLine = StringReplace($sCMDLine, "--single-argument ", "")
				If FileExists($sCMDLine) Then $sCMDLine = '"' & $sCMDLine & '"'
				_SafeRun($sEdge, $sCMDLine)
				If Not _IsPriviledgedInstall() Then Sleep(1000)
			EndIf
		#ce
	
		; Do Either (Run Another App or Run Edge)
		Case StringInStr($sCMDLine, "--app-id") ; "Apps"
			Select
				Case StringInStr($sCMDLine, "--app-fallback-url=") And _GetSettingValue("NoApps", "Bool"); Windows Store "Apps"
					$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)(.*)(--app-fallback-url=)", "")
					$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)(?= --)(.*)", "")
					If _IsSafeURL($sCMDLine) Then
						ShellExecute($sCMDLine)
					Else
						_Log($hLogs[$URIFailures], "Invalid App URL: " & $sCMDLine & @CRLF)
					EndIf
				Case StringInStr($sCMDLine, "--ip-aumid=") ; Edge "Apps"
					If _IsSafeApp($sCMDLine) Then
						_SafeRun($sEdge, $sCMDLine)
					Else
						_Log($hLogs[$URIFailures], "Invalid App URL: " & $sCMDLine & @CRLF)
					EndIf
				Case Else
					_Log($hLogs[$URIFailures], "Invalid App URL: " & $sCMDLine & @CRLF)
			EndSelect

		Case StringInStr($sCMDLine, "ux=copilot") ; CoPilot
			If _GetSettingValue("NoPilot", "Bool") Then
				ShellExecute("ms-settings:")
			Else
				_SafeRun($sEdge, $sCMDLine)
			EndIf

		; Drop Call	
		Case StringInStr($sCMDLine, "--default-search-provider=?")
			_Log($hLogs[$URIFailures], "Blocked Invalid MS-Settings Call: " & $sCMDLine & @CRLF)
		Case StringInStr($sCMDLine, "profiles_settings")
			_Log($hLogs[$URIFailures], "Skipped Profile Settings URL: " & $sCMDLine & @CRLF)

		; Do Either (Drop Call or Run Edge)
		Case StringInStr($sCMDLine, "bing.com/chat") Or StringInStr($sCMDLine, "bing.com%2Fchat") ; Fix BingAI
			If _GetSettingValue("NoChat", "Bool") Then 
				ContinueCase
			Else
				_SafeRun($sEdge, $sCMDLine)
			EndIf

		; Call Default Browser
		Case StringInStr($sCMDLine, "bing.com/spotlight?spotlightid") ; Fix Windows Spotlight
			$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)spotlight\?spotlightid=[^&]+&", "search?")
			ContinueCase
		Case StringInStr($sCMDLine, "&url=") ; Fix Windows 11 Widgets
			ContinueCase
		Case StringInStr($sCMDLine, "microsoft-edge:")
			$aCMDLine = _CMDLineDecode($sCMDLine)

			For $iLoop = 0 To Ubound($aCMDLine) - 1 Step 1
				If $aCMDLine[$iLoop][0] = "url" Then
					$sURL = $aCMDLine[$iLoop][1]				
					ExitLoop
				EndIf
			Next

			If $sURL = "" Then
				_Log($hLogs[$URIFailures], "Command Line Missing Needed Parameters: " & $sCMDLine & @CRLF)
			Else
				_Log($hLogs[$AppGeneral], "Caught Valid URI Call:" & @CRLF & _ArrayToString($aCMDLine, ": ") & @CRLF)
				If _IsSafeURL($sURL) Then
					$sURL = _ModifyURL($sURL)
					ShellExecute($sURL)
				Else
					_Log($hLogs[$URIFailures], "Invalid URL: " & $sCMDLine & @CRLF)
				EndIf
			EndIf

		; Catch Misc Edge Flags (MUST BE LOWEST PRIORITY)
		Case StringLeft($sCMDLine, 2) = "--"
			If _GetSettingValue("RunUnsafe") Then
				_SafeRun($sEdge, $sCMDLine)
			Else
				_Log($hLogs[$AppSecurity], "" & "Blocked Unsafe Flag: " & $sCMDLine & @CRLF)
			EndIf

		; Catch Everything Else
		Case Else
			$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)(.*) microsoft-edge:[\/]*", "") ; Legacy Installs
			$sCMDLine = StringReplace($sCMDLine, "?url=", "")
			If StringInStr($sCMDLine, "%2F") Then $sCMDLine = _WinAPI_UrlUnescape($sCMDLine)
			_Log($hLogs[$AppGeneral], "Caught Other Edge Call:" & @CRLF & $sCMDLine & @CRLF)
			If _IsSafeURL($sCMDLine) Then
				$sCMDLine = _ModifyURL($sCMDLine)
				ShellExecute($sCMDLine)
			Else
				_Log($hLogs[$URIFailures], "Invalid URL: " & $sCMDLine & @CRLF)
			EndIf
	EndSelect
EndFunc

Func _GetDefaultBrowser()

	Local $sProg
	Local Static $sBrowser

	If $sBrowser <> "" Then
		;;;
	Else
		$sProg = RegRead("HKCU\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice", "ProgID")
		$sBrowser = RegRead("HKCR\" & $sProg & "\shell\open\command", "")
		$sBrowser = StringReplace($sBrowser, "%1", "")
	EndIf

	Return $sBrowser

EndFunc
