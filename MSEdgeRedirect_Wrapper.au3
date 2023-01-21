#include-once

#include <Misc.au3>
#include <Array.au3>
#include <String.au3>
#include <GuiComboBox.au3>
#include <EditConstants.au3>
#include <FileConstants.au3>
#include <FontConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <AutoItConstants.au3>
#include <WindowsConstants.au3>

#include "Includes\_Logging.au3"
#include "Includes\_Theming.au3"
#include "Includes\_Settings.au3"
#include "Includes\_Translation.au3"

#include "Includes\TaskScheduler.au3"

; TODO: Why have <Setting>PATH values for Custom handlers... Rewrite that.

Global $sVersion
Global $bIsPriv = _IsPriviledgedInstall()
Global Enum $bNoApps, $bNoBing, $bNoImgs, $bNoMSN, $bNoNews, $bNoPDFs, $bNoTray, $bNoUpdates, $sImages, $sImagePath, $sNews, $sPDFApp, $sSearch, $sSearchPath, $sStartMenu, $bStartup, $sWeather, $sWeatherPath

If @Compiled Then
	$sVersion = FileGetVersion(@ScriptFullPath)
Else
	$sVersion = "x.x.x.x"
EndIf

Func RunInstall(ByRef $aConfig, ByRef $aSettings, $bSilent = False)

	Local $sArgs = ""
	Local Enum $bManaged = 1, $vMode

	SetOptionsRegistry("NoApps"     , $aSettings[$bNoApps]     , $aConfig)
	SetOptionsRegistry("NoBing"     , $aSettings[$bNoBing]     , $aConfig)
	SetOptionsRegistry("NoImgs"     , $aSettings[$bNoImgs]     , $aConfig)
	SetOptionsRegistry("NoMSN"      , $aSettings[$bNoMSN]      , $aConfig)
	SetOptionsRegistry("NoNews"     , $aSettings[$bNoNews]     , $aConfig)
	SetOptionsRegistry("NoPDFs"     , $aSettings[$bNoPDFs]     , $aConfig)
	SetOptionsRegistry("NoTray"     , $aSettings[$bNoTray]     , $aConfig)
	SetOptionsRegistry("NoUpdates"  , $aSettings[$bNoUpdates]  , $aConfig)
	SetOptionsRegistry("Images"     , $aSettings[$sImages]     , $aConfig)
	SetOptionsRegistry("ImagePath"  , $aSettings[$sImagePath]  , $aConfig)
	SetOptionsRegistry("News"       , $aSettings[$sNews]       , $aConfig)
	SetOptionsRegistry("PDFApp"     , $aSettings[$sPDFApp]     , $aConfig)
	SetOptionsRegistry("Search"     , $aSettings[$sSearch]     , $aConfig)
	SetOptionsRegistry("SearchPath" , $aSettings[$sSearchPath] , $aConfig)
	SetOptionsRegistry("Weather"    , $aSettings[$sWeather]    , $aConfig)
	SetOptionsRegistry("WeatherPath", $aSettings[$sWeatherPath], $aConfig)

	If $aConfig[$vMode] Then
		If Not FileCopy(@ScriptFullPath, $sDrive & "\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE) Then
			FileWrite($hLogs[$AppFailures], _NowCalc() & " - [CRITICAL] Unable to copy application to " & $sDrive & "'\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe'" & @CRLF)
			If Not $bSilent Then
				MsgBox($MB_ICONERROR+$MB_OK, _
					"[CRITICAL]", _
					"Unable to copy application to " & $sDrive & "'\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe'")
			EndIf
			Exit 29 ; ERROR_WRITE_FAULT
		EndIf
	Else
		If $aSettings[$bNoTray] Then $sArgs = "/hide"
		If Not FileCopy(@ScriptFullPath, @LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE) Then
			FileWrite($hLogs[$AppFailures], _NowCalc() & " - [CRITICAL] Unable to copy application to '" & @LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe'" & @CRLF)
			If Not $bSilent Then
				MsgBox($MB_ICONERROR+$MB_OK, _
					"[CRITICAL]", _
					"Unable to copy application to '" & @LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe'")
			EndIf
			Exit 29 ; ERROR_WRITE_FAULT
		EndIf
		If $aSettings[$bStartup] Then
			If Not FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @StartupDir & "\MSEdgeRedirect.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", $sArgs) Then
				FileWrite($hLogs[$AppFailures], _NowCalc() & " - [WARNING] Unable to create application link in '" & @StartupDir & "\MSEdgeRedirect.lnk'" & @CRLF)
			EndIf
		EndIf
	EndIf

EndFunc

Func RunPDFCheck($bSilent = False)

	If StringRegExp(RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice", "ProgId"), "(?i)(ms|microsoft)edge.*") Then
		If Not $bSilent Then
			MsgBox($MB_ICONERROR+$MB_OK, _
				"Edge Set As Default PDF Handler", _
				"You must set a different Default PDF Handler to use this feature!")
		EndIf
		Return False
	EndIf
	Return True

EndFunc

Func RunRemoval($bUpdate = False)

	Local $hTS
	Local $aPIDs
	Local $sHive = ""
	Local $sLocation = ""

	$aPIDs = ProcessList("msedgeredirect.exe")
	For $iLoop = 1 To $aPIDs[0][0] Step 1
		If $aPIDs[$iLoop][1] <> @AutoItPID Then ProcessClose($aPIDs[$iLoop][1])
	Next

	If $bUpdate Then
		$sHive = _IsInstalled()[1]
	Else
		If $bIsAdmin Then
			$sHive = "HKLM"
		Else
			$sHive = "HKCU"
		EndIf
	EndIf

	If $sHive = "HKLM" Then
		$sLocation = $sDrive & "\Program Files\MSEdgeRedirect\"
	ElseIf $sHive = "HKCU" THen
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
	Else
		FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Failed to Determine Registry Hive for Uninstall." & @CRLF)
		FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "DEBUG: " & _ArrayToString(_IsInstalled()) &  @CRLF)
		Exit 1359 ; ERROR_INTERNAL_ERROR
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

	; Start Menu Shortcuts
	FileDelete(@StartupDir & "\MSEdgeRedirect.lnk")
	DirRemove(@ProgramsCommonDir & "\MSEdgeRedirect", $DIR_REMOVE)
	DirRemove(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect", $DIR_REMOVE)

	; Parent Registry Key
	RegEnumKey($sHive & "\SOFTWARE\Robert Maehl Software", 1)
	If @error Then RegDelete($sHive & "\SOFTWARE\Robert Maehl Software")

	If $bIsAdmin Then
		For $iLoop = 1 To $aEdges[0] Step 1
			If $iLoop = $aEdges[0] Then ExitLoop ; Skip ie_to_edge_stub
			If FileExists(StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe")) Then ; Pre-0.7.3.0
				FileDelete(StringReplace($aEdges[$iLoop], "msedge.exe", "msedge_no_ifeo.exe"))
			EndIf
			If FileExists(StringReplace($aEdges[$iLoop], "Application\msedge.exe", "IFEO\")) Then ; 0.7.3.0+
				DirRemove(StringReplace($aEdges[$iLoop], "Application\msedge.exe", "IFEO\"))
			EndIf
		Next
		$hTS = _TS_Open() ; 0.7.2.0
		_TS_TaskDelete($hTS, "\MSEdgeRedirect\Update Edge.xml")
		_TS_TaskDelete($hTS, "\MSEdgeRedirect\Update Edge Beta.xml")
		_TS_TaskDelete($hTS, "\MSEdgeRedirect\Update Edge Canary.xml")
		_TS_TaskDelete($hTS, "\MSEdgeRedirect\Update Edge Dev.xml")
		_TS_FolderDelete($hTS, "\MSEdgeRedirect")
		_TS_Close($hTS)
	EndIf

	If $bUpdate Then
		FileDelete($sLocation & "*")
	Else
		Run(@ComSpec & " /c " & 'ping google.com && del /Q "' & $sLocation & '*"', "", @SW_HIDE)
		Exit
	EndIf

EndFunc

Func RunRepair()

	If $bIsAdmin Then
		For $iLoop = 1 To $aEdges[0] Step 1
			RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop, "Debugger")
			If @error Then
				;;;
			Else
				If $iLoop = $aEdges[0] Then
					;;;
				Else
					FileCreateNTFSLink(StringReplace($aEdges[$iLoop], "\msedge.exe", ""), StringReplace($aEdges[$iLoop], "Application\msedge.exe", "IFEO\"), $FC_OVERWRITE)
				EndIf
			EndIf
		Next
		Exit
	Else
		Exit 5 ; ERROR_ACCESS_DENIED
	EndIf

EndFunc

Func RunSetup($bUpdate = False, $bSilent = False, $iPage = 0, $hSetupFile = @ScriptDir & "\Setup.ini")

	Local $hMsg
	Local $sArgs = ""
	Local $iMode = $iPage
	Local $sEdges
	Local $sEngine
	Local $sImgEng
	Local $sHandler
	Local $hChannels[5]
	Local $aChannels[5] = [True, False, False, False, True]
	Local $sWeatherEng

	Local $aConfig[3] = [$hSetupFile, False, "Service"] ; Default Setup.ini Values
	Local Enum $hFile, $bManaged, $vMode

	Local $aSettings[18] = [False, False, False, False, False, False, False, False, "", "", "", "", "", "", "Full", True, "", ""]

	If $bSilent Then

		If $bUpdate Then
			$aSettings[$bNoApps] = _Bool(_GetSettingValue("NoApps"))
			$aSettings[$bNoBing] = _Bool(_GetSettingValue("NoBing"))
			$aSettings[$bNoImgs] = _Bool(_GetSettingValue("NoImgs"))
			$aSettings[$bNoMSN] = _Bool(_GetSettingValue("NoMSN"))
			$aSettings[$bNoNews] = _Bool(_GetSettingValue("NoNews"))
			$aSettings[$bNoPDFs] = _Bool(_GetSettingValue("NoPDFs"))
			$aSettings[$bNoTray] = _Bool(_GetSettingValue("NoTray"))
			$aSettings[$bNoUpdates] = _Bool(_GetSettingValue("NoUpdates"))
			If $aSettings[$bNoBing] Then
				$aSettings[$sSearch] = _GetSettingValue("Search")
				$aSettings[$sSearchPath] = _GetSettingValue("SearchPath")
			EndIf
			If $aSettings[$bNoImgs] Then
				$aSettings[$sImages] = _GetSettingValue("Images")
				$aSettings[$sImagePath] = _GetSettingValue("ImagePath")
			EndIf
			If $aSettings[$bNoPDFs] Then $aSettings[$sPDFApp] = _GetSettingValue("PDFApp")
			If $aSettings[$bNoMSN] Then
				$aSettings[$sWeather] = _GetSettingValue("Weather")
				$aSettings[$sWeatherPath] = _GetSettingValue("WeatherPath")
			EndIf
			If $aSettings[$bNoNews] Then $aSettings[$sNews] = _GetSettingValue("News")
		EndIf

		If $aConfig[$hFile] = "WINGET" Then
			$aConfig[$vMode] = $bIsAdmin
			; Bypass file checks, IniReads, use default values
		ElseIf Not FileExists($aConfig[$hFile]) And Not $bUpdate Then
			Exit 2 ; ERROR_FILE_NOT_FOUND
		Else
			$aConfig[$bManaged] = _Bool(IniRead($aConfig[$hFile], "Config", "Managed", False))
			$aConfig[$vMode] = IniRead($aConfig[$hFile], "Config", "Mode", "Service")

			If $aConfig[$vMode] = "active" Then
				$aConfig[$vMode] = True
			Else
				$aConfig[$vMode] = False
			EndIf

			; TODO: Merge with _GetSettingValue(Value, Forced Location)
			$aSettings[$bNoApps] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoApps", $aSettings[$bNoApps]))
			$aSettings[$bNoBing] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoBing", $aSettings[$bNoBing]))
			$aSettings[$bNoImgs] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoImgs", $aSettings[$bNoImgs]))
			$aSettings[$bNoMSN] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoMSN", $aSettings[$bNoMSN]))
			$aSettings[$bNoNews] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoNews", $aSettings[$bNoNews]))
			$aSettings[$bNoPDFs] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoPDFs", $aSettings[$bNoPDFs]))
			$aSettings[$bNoTray] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoTray", $aSettings[$bNoTray]))
			$aSettings[$bNoUpdates] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoUpdates", $aSettings[$bNoUpdates]))
			$aSettings[$sImages] = _Bool(IniRead($aConfig[$hFile], "Settings", "Images", $aSettings[$sImages]))
			$aSettings[$sImagePath] = _Bool(IniRead($aConfig[$hFile], "Settings", "ImagePath", $aSettings[$sImagePath]))
			$aSettings[$sNews] = _Bool(IniRead($aConfig[$hFile], "Settings", "News", $aSettings[$sNews]))
			$aSettings[$sPDFApp] = IniRead($aConfig[$hFile], "Settings", "PDFApp", $aSettings[$sPDFApp])
			$aSettings[$sSearchPath] = IniRead($aConfig[$hFile], "Settings", "SearchPath", $aSettings[$sSearchPath])
			$aSettings[$sStartMenu] = IniRead($aConfig[$hFile], "Settings", "StartMenu", $aSettings[$sStartMenu])
			$aSettings[$sSearch] = IniRead($aConfig[$hFile], "Settings", "Search", $aSettings[$sSearch])
			$aSettings[$bStartup] = _Bool(IniRead($aConfig[$hFile], "Settings", "Startup", $aSettings[$bStartup]))
			$aSettings[$sWeather] = IniRead($aConfig[$hFile], "Settings", "Weather", $aSettings[$sWeather])
			$aSettings[$sWeatherPath] = IniRead($aConfig[$hFile], "Settings", "WeatherPath", $aSettings[$sWeatherPath])

			$sEdges = IniRead($aConfig[$hFile], "Settings", "Edges", "")
			If StringInStr($sEdges, "Stable") Then $aChannels[0] = True
			If StringInStr($sEdges, "Beta") Then $aChannels[1] = True
			If StringInStr($sEdges, "Dev") Then $aChannels[2] = True
			If StringInStr($sEdges, "Canary") Then $aChannels[3] = True
			If StringInStr($sEdges, "Removed") Then $aChannels[4] = True

		EndIf

		If ($aConfig[$bManaged] Or $aConfig[$vMode]) And Not $bIsAdmin Then
			If $aConfig[$hFile] = "WINGET" Then
				FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Failed to Self Escalate for Deployment." & @CRLF)
			Else
				Exit 5 ; ERROR_ACCESS_DENIED
			EndIf
		EndIf

		For $iLoop = 0 To UBound($aChannels) - 1 Step 1
			If $aChannels[$iLoop] = True Then ExitLoop
			If $iLoop = UBound($aChannels) - 1 Then
				If $aConfig[$hFile] = "WINGET" Then
					FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Failed to Self Validate IEFO Channels." & @CRLF)
				Else
					Exit 1359 ; ERROR_INTERNAL_ERROR
				EndIf
			EndIf
		Next

		For $iLoop = $bNoApps To $bNoUpdates Step 1
			If Not IsBool($aSettings[$iLoop]) Then
				If $aConfig[$hFile] = "WINGET" Then
					FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Failed to Self Validate Boolean Settings." & @CRLF)
				Else
					Exit 1359 ; ERROR_INTERNAL_ERROR
				EndIf
			EndIf
		Next
		If Not IsBool($aSettings[$bStartup]) Then
			If $aConfig[$hFile] = "WINGET" Then
				FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Failed to Self Validate Startup Boolean." & @CRLF)
			Else
				Exit 1359 ; ERROR_INTERNAL_ERROR
			EndIf
		EndIf

		If $bUpdate Then RunRemoval(True)
		RunInstall($aConfig, $aSettings, $bSilent)
		SetAppRegistry($aConfig)
		SetAppShortcuts($aConfig, $aSettings)
		If $aConfig[$vMode] Then
			SetIFEORegistry($aChannels)
		Else
			If $aSettings[$bNoTray] Then $sArgs = "/hide"
			ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $sArgs, @LocalAppDataDir & "\MSEdgeRedirect\")
		EndIf
		Exit

	Else

		Local $aPages[4]
		Local Enum $hLicense, $hMode, $hSettings, $hFinish, $hExit

		If @Compiled And Not _GetSettingValue("NoUpdates") Then RunUpdateCheck()

		If StringInStr($bUpdate, "HKLM") And Not $bIsAdmin And Not @Compiled Then
			MsgBox($MB_ICONERROR+$MB_OK, _
				"Admin Required", _
				"Unable to update an Admin Install without Admin Rights!")
			FileWrite($hLogs[$AppFailures], _NowCalc() & " - " & "Non Admin Update Attempt on Admin Install. EXITING!" & @CRLF)
			For $iLoop = 0 To UBound($hLogs) - 1
				FileClose($hLogs[$iLoop])
			Next
			Exit 5 ; ERROR_ACCESS_DENIED
		EndIf

		; Disable Scaling
		If @OSVersion = 'WIN_10' Then DllCall(@SystemDir & "\User32.dll", "bool", "SetProcessDpiAwarenessContext", "HWND", "DPI_AWARENESS_CONTEXT" - 1)

		Local $hInstallGUI = GUICreate("MSEdgeRedirect " & $sVersion & " Setup", 640, 480)

		GUICtrlCreateLabel("", 0, 0, 180, 420)
		GUICtrlSetBkColor(-1, 0x00A4EF)

		GUICtrlCreateIcon("", -1, 26, 26, 128, 128)
		If @Compiled Then
			_SetBkSelfIcon(-1, "", 0x00A4EF, @ScriptFullPath, 201, 128, 128)
		Else
			_SetBkIcon(-1, "", 0x00A4EF, @ScriptDir & "\assets\MSEdgeRedirect.ico", -1, 128, 128)
		EndIf

		Local $hHelp = GUICtrlCreateButton("Help", 20, 435, 90, 30)

		Local $hBack = GUICtrlCreateButton("< Back", 330, 435, 90, 30)
		GUICtrlSetState(-1, $GUI_DISABLE)
		Local $hNext = GUICtrlCreateButton("Next >", 420, 435, 90, 30)
		If $iPage = $hLicense Then
			GUICtrlSetState(-1, $GUI_DISABLE)
		ElseIf $iPage = $hSettings Then
			GUICtrlSetData(-1, "Save")
		EndIf
		Local $hCancel = GUICtrlCreateButton("Cancel", 530, 435, 90, 30)

		#Region License Page
		$aPages[$hLicense] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)
		FileInstall("./LICENSE", @LocalAppDataDir & "\MSEdgeRedirect\License.txt")

		If $bUpdate Then
			GUICtrlCreateLabel("Please read the following License. You must accept the terms of the license before continuing with the upgrade.", 20, 20, 420, 40)
		Else
			GUICtrlCreateLabel("Please read the following License. You must accept the terms of the license before continuing with the installation.", 20, 20, 420, 40)
		EndIf

		GUICtrlCreateEdit("TL;DR: It's FOSS, you can edit it, repackage it, eat it (not recommended), or throw it at your neighbor Steve (depends on the Steve), but changes to it must be LGPL v3 too." & _
			@CRLF & @CRLF & _
			FileRead(@LocalAppDataDir & "\MSEdgeRedirect\License.txt"), 20, 60, 420, 280, $ES_READONLY + $WS_VSCROLL)

		Local $hAgree = GUICtrlCreateRadio("I accept this license", 20, 350, 420, 20)
		Local $hDisagree = GUICtrlCreateRadio("I don't accept this license", 20, 370, 420, 20)
		GUICtrlSetState(-1, $GUI_CHECKED)

		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Mode Page
		$aPages[$hMode] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)
		If $bUpdate Then
			GUICtrlCreateLabel("MSEdgeRedirect " & $sVersion & " Update", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
;			GUICtrlCreateLabel("Click Next to continue the Update of MSEdgeRedirect after customizing your preferred mode", 20, 40, 420, 40)
;			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Install MSEdgeRedirect " & $sVersion, 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
;			GUICtrlCreateLabel("Click Next to continue the Install of MSEdgeRedirect after customizing your preferred mode", 20, 40, 420, 40)
;			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		GUICtrlCreateGroup("Mode", 20, 60, 420, 340)
			Local $hService = GUICtrlCreateRadio("Service Mode" & @CRLF & _
				@CRLF & _
				"* Single User Install" & @CRLF & _
				"* Less System Intrusive" & @CRLF & _
				"* Less AV false positives" & @CRLF & _
				"* Doesn't require Admin Rights" & @CRLF & _
				@CRLF & _
				"MSEdgeRedirect stays running in the background. Detected Edge data is redirected to your default browser.", _
				50, 80, 380, 130, $BS_TOP+$BS_MULTILINE)
			If Not $bIsAdmin Then GUICtrlSetState(-1, $GUI_CHECKED)

			GUICtrlCreateIcon("imageres.dll", 78, 30, 230, 16, 16)
			Local $hActive = GUICtrlCreateRadio("Active Mode - RECOMMENDED" & @CRLF & _
				@CRLF & _
				"* Better Performance" & @CRLF & _
				"* System Wide Install" & @CRLF & _
				"* Finer Redirection Control" & @CRLF & _
				"* No Startup or Tray Icon Needed" & @CRLF & _
				"* Supports AveYo's Edge Remover Tool" & @CRLF & _
				@CRLF & _
				"MSEdgeRedirect only runs when a selected Edge is launched, similary to the old EdgeDeflector app.", _
				50, 230, 380, 130, $BS_TOP+$BS_MULTILINE)
			If $bIsAdmin Then GUICtrlSetState(-1, $GUI_CHECKED)

		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Settings Page
		$aPages[$hSettings] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)

		If $bUpdate Then
			GUICtrlCreateLabel("MSEdgeRedirect " & $sVersion & " Update", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
;			GUICtrlCreateLabel("Click Install to continue the Update of MSEdgeRedirect after customizing your preferred settings", 20, 40, 420, 40)
;			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Install MSEdgeRedirect " & $sVersion, 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
;			GUICtrlCreateLabel("Click Install to continue the Install of MSEdgeRedirect after customizing your preferred settings", 20, 40, 420, 40)
;			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		GUICtrlCreateGroup("Active Mode Options", 20, 60, 420, 70)
			$hChannels[0] = GUICtrlCreateCheckbox("Edge Stable", 50, 80, 95, 20)
			$hChannels[1] = GUICtrlCreateCheckbox("Edge Beta", 145, 80, 95, 20)
			$hChannels[2] = GUICtrlCreateCheckbox("Edge Dev", 240, 80, 95, 20)
			$hChannels[3] = GUICtrlCreateCheckbox("Edge Canary", 335, 80, 95, 20)
			$hChannels[4] = GUICtrlCreateCheckbox("Edge Removed Using AveYo's Edge Remover (Auto Detected)", 50, 100, 380, 20)
			GUICtrlSetState(-1, $GUI_DISABLE)

			Select
				Case Not $bIsAdmin And @Compiled
					GUICtrlSetState($hChannels[0], $GUI_DISABLE)
					GUICtrlSetState($hChannels[1], $GUI_DISABLE)
					GUICtrlSetState($hChannels[2], $GUI_DISABLE)
					GUICtrlSetState($hChannels[3], $GUI_DISABLE)
				Case FileExists($aEdges[5])
					GUICtrlSetState($hChannels[4], $GUI_CHECKED)
					ContinueCase
				Case Else
					If $bUpdate Or $iMode = $hSettings Then
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER1", "Debugger") Then GUICtrlSetState($hChannels[0], $GUI_CHECKED)
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER2", "Debugger") Then GUICtrlSetState($hChannels[1], $GUI_CHECKED)
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER3", "Debugger") Then GUICtrlSetState($hChannels[2], $GUI_CHECKED)
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER4", "Debugger") Then GUICtrlSetState($hChannels[3], $GUI_CHECKED)
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER5", "Debugger") Then GUICtrlSetState($hChannels[4], $GUI_CHECKED)
					Else
						GUICtrlSetState($hChannels[0], $GUI_CHECKED)
					EndIf
			EndSelect

		GUICtrlCreateGroup("Service Mode Options", 20, 140, 420, 70)
			Local $hNoIcon = GUICtrlCreateCheckbox("Hide Service Mode from Tray", 50, 160, 380, 20)
			Local $hStartup = GUICtrlCreateCheckbox("Start Service Mode With Windows", 50, 180, 380, 20)

			If $bIsAdmin Then
				GUICtrlSetState($hStartup, $GUI_DISABLE)
				GUICtrlSetState($hNoIcon, $GUI_DISABLE)
			ElseIf $bUpdate Then
				GUICtrlSetState($hStartup, FileExists(@StartupDir & "\MSEdgeRedirect.lnk"))
				GUICtrlSetState($hNoIcon, _GetSettingValue("NoApps"))
			EndIf

		GUICtrlCreateGroup("Additional Redirections", 20, 220, 420, 190)
			Local $hSearch = GUICtrlCreateCheckbox("Bing Search:", 50, 240, 180, 20)
			Local $hEngine = GUICtrlCreateCombo("", 50, 260, 180, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "Ask|Baidu|Custom|DuckDuckGo|Ecosia|Google|Sogou|Yahoo|Yandex", "Google")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoImgs = GUICtrlCreateCheckbox("Bing Images:", 240, 240, 180, 20)
			Local $hImgSRC = GUICtrlCreateCombo("", 240, 260, 180, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "Baidu|Custom|DuckDuckGo|Ecosia|Google|Sogou|Yahoo|Yandex", "Google")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoNews = GUICtrlCreateCheckbox("MSN News: (ALPHA)", 50, 285, 180, 20)
			Local $hNewSRC = GUICtrlCreateCombo("", 50, 305, 180, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "DuckDuckGo|Google", "Google")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoMSN = GUICtrlCreateCheckbox("MSN Weather:", 240, 285, 180, 20)
			Local $hWeather = GUICtrlCreateCombo("", 240, 305, 180, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "AccuWeather|Custom|DarkSky|Weather.com|Weather.gov|Windy|WUnderground|Ventusky|Yandex", "Weather.com")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoPDFs = GUICtrlCreateCheckbox("PDF Viewer:", 50, 330, 180, 20)
			Local $hPDFSrc = GUICtrlCreateCombo("", 50, 350, 180, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "Default|Custom", "Default")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoApps = GUICtrlCreateCheckbox("Windows Store 'Apps'", 50, 375, 180, 20)

		If $bUpdate Then
			GUICtrlSetState($hNoApps, _GetSettingValue("NoApps"))
			GUICtrlSetState($hSearch, _GetSettingValue("NoBing"))
			If _IsChecked($hSearch) Then
				GUICtrlSetState($hEngine, $GUI_ENABLE)
				GUICtrlSetData($hEngine, _GetSettingValue("Search"))
				$sEngine = _GetSettingValue("SearchPath")
			EndIf
			GUICtrlSetState($hNoImgs, _GetSettingValue("NoImgs"))
			If _IsChecked($hNoImgs) Then
				GUICtrlSetState($hImgSRC, $GUI_ENABLE)
				GUICtrlSetData($hImgSRC, _GetSettingValue("Images"))
				$sImgEng = _GetSettingValue("ImagePath")
			EndIf
			GUICtrlSetState($hNoMSN, _GetSettingValue("NoMSN"))
			If _IsChecked($hNoMSN) Then
				GUICtrlSetState($hWeather, $GUI_ENABLE)
				GUICtrlSetData($hWeather, _GetSettingValue("Weather"))
				$sWeatherEng = _GetSettingValue("WeatherPath")
			EndIf
			GUICtrlSetState($hNoNews, _GetSettingValue("NoNews"))
			If _IsChecked($hNoNews) Then
				GUICtrlSetState($hNewSRC, $GUI_ENABLE)
				GUICtrlSetData($hNewSRC, _GetSettingValue("News"))
			EndIf
			GUICtrlSetState($hNoPDFs, _GetSettingValue("NoPDFs"))
			If _IsChecked($hNoPDFs) Then
				GUICtrlSetState($hPDFSrc, $GUI_ENABLE)
				If StringInStr(_GUICtrlComboBox_GetList($hPDFSrc), _GetSettingValue("PDFApp")) Then
					GUICtrlSetData($hPDFSrc, _GetSettingValue("PDFApp"))
				Else
					GUICtrlSetData($hPDFSrc, "Custom")
				EndIf
				$sHandler = _GetSettingValue("PDFApp")
			EndIf
		EndIf

		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Finish Page
		$aPages[$hFinish] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)

		If $bUpdate Then
			GUICtrlCreateLabel("Updated Successfully", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Installed Successfully", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		Local $hLaunch = GUICtrlCreateCheckbox("Launch Service Mode Now", 20, 200, 190, 20)
		Local $hAppLnk = GUICtrlCreateCheckbox("Create Start Menu Shortcuts", 20, 220, 190, 20)
		GUICtrlSetState(-1, $GUI_CHECKED)
		Local $hDonate = GUICtrlCreateCheckbox("Donate to the Project via PayPal", 20, 240, 190, 20)
		Local $hHelpUs = GUICtrlCreateCheckbox("Help us get off Google's Blacklist", 20, 260, 190, 20)

		GUISwitch($hInstallGUI)
		#EndRegion

		GUISetState(@SW_SHOW, $hInstallGUI)
		GUISetState(@SW_SHOW, $aPages[$iPage])

		While True
			$hMsg = GUIGetMsg()

			Select

				Case $hMsg = $GUI_EVENT_CLOSE or $hMsg = $hCancel
					Exit

				Case $hMsg = $hAgree or $hMsg = $hDisagree
					If _IsChecked($hAgree) Then
						GUICtrlSetState($hNext, $GUI_ENABLE)
					Else
						GUICtrlSetState($hNext, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hHelp
					Switch $iPage
						Case $hLicense
							ShellExecute("https://msedgeredirect.com/wiki/Installer-Overview#license-page")
						Case $hMode
							ShellExecute("https://msedgeredirect.com/wiki/Installer-Overview#mode-page")
						Case $hSettings
							ShellExecute("https://msedgeredirect.com/wiki/Installer-Overview#settings-page")
						Case $hFinish
							ShellExecute("https://msedgeredirect.com/wiki/Installer-Overview#finish-page")
						Case Else

					EndSwitch

				Case $hMsg = $hBack
					Switch $iPage - 1
						Case $hLicense
							GUICtrlSetState($hBack, $GUI_DISABLE)
							GUICtrlSetState($hNext, $GUI_ENABLE)
						Case $hMode
							GUICtrlSetData($hNext, "Next >")
						Case $hSettings
							If $bUpdate Then
								GUICtrlSetData($hNext, "Update")
							Else
								GUICtrlSetData($hNext, "Install")
							EndIf
					EndSwitch
					GUISetState(@SW_HIDE, $aPages[$iPage])
					GUISetState(@SW_SHOW, $aPages[$iPage - 1])
					$iPage -= 1

				Case $hMsg = $hNext
					Switch $iPage + 1
						Case $hMode
							GUICtrlSetState($hBack, $GUI_ENABLE)
						Case $hSettings
							If $bUpdate Then
								GUICtrlSetData($hNext, "Update")
							Else
								GUICtrlSetData($hNext, "Install")
							EndIf
						Case $hFinish
							# 8.0.0.0 Refactor
							If $bUpdate And $iMode <> $hSettings Then
								RunRemoval(True)
							Else
								FileDelete(@StartupDir & "\MSEdgeRedirect.lnk")
							EndIf

							If $iMode = $hSettings Then
								$aConfig[$vMode] = $bIsAdmin
							Else
								$aConfig[$vMode] = _IsChecked($hActive)
							EndIf

							$aSettings[$bNoApps] = _IsChecked($hNoApps)
							$aSettings[$bNoBing] = _IsChecked($hSearch)
							$aSettings[$bNoImgs] = _IsChecked($hNoImgs)
							$aSettings[$bNoMSN] = _IsChecked($hNoMSN)
							$aSettings[$bNoNews] = _IsChecked($hNoNews)
							$aSettings[$bNoPDFs] = _IsChecked($hNoPDFs)
							$aSettings[$bNoTray] = _IsChecked($hNoIcon)
							$aSettings[$sImages] = GUICtrlRead($hImgSRC)
							$aSettings[$sImagePath] = $sImgEng
							$aSettings[$sNews] = GUICtrlRead($hNewSRC)
							$aSettings[$sPDFApp] = $sHandler
							$aSettings[$sSearch] = GUICtrlRead($hEngine)
							$aSettings[$sSearchPath] = $sEngine
							$aSettings[$bStartup] = _IsChecked($hStartup)
							$aSettings[$sWeather] = GUICtrlRead($hWeather)
							$aSettings[$sWeatherPath] = $sWeatherEng

							GUISetState(@SW_HIDE, $hSettings)
							RunInstall($aConfig, $aSettings)
							SetAppRegistry($aConfig)
							If $aConfig[$vMode] Then
								For $iLoop = 0 To UBound($aChannels) - 1 Step 1
									$aChannels[$iLoop] = _IsChecked($hChannels[$iLoop])
								Next
								SetIFEORegistry($aChannels)
							EndIf
							If $iMode = $hSettings Then Return
							GUICtrlSetData($hNext, "Finish")
							GUICtrlSetState($hHelp, $GUI_DISABLE)
							GUICtrlSetState($hBack, $GUI_DISABLE)
							GUICtrlSetState($hCancel, $GUI_DISABLE)
							If _IsChecked($hActive) Then
								GUICtrlSetState($hLaunch, $GUI_DISABLE)
							Else
								GUICtrlSetState($hLaunch, $GUI_CHECKED)
							EndIf
						Case $hExit
							If _IsChecked($hAppLnk) Then SetAppShortcuts($aConfig, $aSettings)
							If _IsChecked($hDonate) Then ShellExecute("https://paypal.me/rhsky")
							If _IsChecked($hHelpUs) Then ShellExecute("https://safebrowsing.google.com/safebrowsing/report_error/?url=https://github.com/rcmaehl/MSEdgeRedirect")
							If Not $aConfig[$vMode] And _IsChecked($hLaunch) Then
								If $aSettings[$bNoTray] Then $sArgs = "/hide"
								ShellExecute(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $sArgs, @LocalAppDataDir & "\MSEdgeRedirect\")
							EndIf
							Exit
					EndSwitch
					GUISetState(@SW_HIDE, $aPages[$iPage])
					GUISetState(@SW_SHOW, $aPages[$iPage + 1])
					$iPage += 1

				Case $hMsg = $hActive or $hMsg = $hService
					If @Compiled And _IsChecked($hActive) And Not $bIsAdmin Then
						If ShellExecute(@ScriptFullPath, "", @ScriptDir, "RunAs") Then Exit
						GUICtrlSetState($hActive, $GUI_UNCHECKED)
						GUICtrlSetState($hService, $GUI_CHECKED)
						MsgBox($MB_ICONERROR+$MB_OK, _
							"Admin Required", _
							"Unable to install Active Mode without Admin Rights!")
						FileWrite($hLogs[$AppFailures], _NowCalc() & " - " & "Active Mode UAC Elevation Attempt Failed!" & @CRLF)
					EndIf
					If _IsChecked($hService) Then
						;GUICtrlSetState($hInstall, $GUI_ENABLE)
						GUICtrlSetState($hStartup, $GUI_ENABLE)
						GUICtrlSetState($hNoIcon, $GUI_ENABLE)
						GUICtrlSetState($hChannels[0], $GUI_DISABLE)
						GUICtrlSetState($hChannels[1], $GUI_DISABLE)
						GUICtrlSetState($hChannels[2], $GUI_DISABLE)
						GUICtrlSetState($hChannels[3], $GUI_DISABLE)
					Else
						GUICtrlSetState($hStartup, $GUI_DISABLE)
						GUICtrlSetState($hNoIcon, $GUI_DISABLE)
						If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe\0", "Debugger") Then
							GUICtrlSetState($hChannels[0], $GUI_DISABLE)
							GUICtrlSetState($hChannels[1], $GUI_DISABLE)
							GUICtrlSetState($hChannels[2], $GUI_DISABLE)
							GUICtrlSetState($hChannels[3], $GUI_DISABLE)
							GUICtrlSetState($hChannels[4], $GUI_DISABLE)
							GUICtrlSetState($hChannels[4], $GUI_CHECKED)
						 Else
							If $bUpdate Or $iMode = $hSettings Then
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER1", "Debugger") Then GUICtrlSetState($hChannels[0], $GUI_CHECKED)
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER2", "Debugger") Then GUICtrlSetState($hChannels[1], $GUI_CHECKED)
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER3", "Debugger") Then GUICtrlSetState($hChannels[2], $GUI_CHECKED)
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER4", "Debugger") Then GUICtrlSetState($hChannels[3], $GUI_CHECKED)
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER5", "Debugger") Then GUICtrlSetState($hChannels[4], $GUI_CHECKED)
								If RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\0", "Debugger") Then GUICtrlSetState($hChannels[4], $GUI_CHECKED)
							Else
								GUICtrlSetState($hChannels[0], $GUI_ENABLE)
								GUICtrlSetState($hChannels[1], $GUI_ENABLE)
								GUICtrlSetState($hChannels[2], $GUI_ENABLE)
								GUICtrlSetState($hChannels[3], $GUI_ENABLE)
								GUICtrlSetState($hChannels[0], $GUI_CHECKED)
							EndIf
						EndIf
						ContinueCase
					EndIf

				Case $hMsg = $hChannels[0] Or $hMsg = $hChannels[1] Or $hMsg = $hChannels[2] Or $hMsg = $hChannels[3]
					;GUICtrlSetState($hInstall, $GUI_DISABLE)
					For $iLoop = 0 To UBound($aChannels) - 1 Step 1
						If _IsChecked($hChannels[$iLoop]) Then
							;GUICtrlSetState($hInstall, $GUI_ENABLE)
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
					$sEngine = InputBox("Enter Search Engine URL", "Enter the URL format of the custom Search Engine to use", "https://duckduckgo.com/?q=")
					If @error Then GUICtrlSetData($hEngine, "Google")

				Case $hMsg = $hNoImgs
					If _IsChecked($hNoImgs) Then
						GUICtrlSetState($hImgSRC, $GUI_ENABLE)
					Else
						GUICtrlSetState($hImgSRC, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hImgSRC And GUICtrlRead($hImgSRC) = "Custom"
					$sImgEng = InputBox("Enter Image Search Engine URL", "Enter the URL format of the custom Image Search Engine to use", "https://duckduckgo.com/?ia=images&iax=images&q=")
					If @error Then GUICtrlSetData($hImgSRC, "Google")

				Case $hMsg = $hNoMSN
					If _IsChecked($hNoMSN) Then
						GUICtrlSetState($hWeather, $GUI_ENABLE)
					Else
						GUICtrlSetState($hWeather, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hWeather And GUICtrlRead($hWeather) = "Custom"
					$sWeatherEng = InputBox("Enter Weather Engine URL", "Enter the URL format of the custom Weather Engine to use (REQUIRES LAT,LONG SUPPORT!)", "https://www.accuweather.com/en/search-locations?query=")
					If @error Then GUICtrlSetData($hImgSRC, "Weather.com")

				Case $hMsg = $hNoNews
					If _IsChecked($hNoNews) Then
						GUICtrlSetState($hNewSRC, $GUI_ENABLE)
					Else
						GUICtrlSetState($hNewSRC, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hNoPDFs
					If _IsChecked($hNoPDFs) Then
						GUICtrlSetState($hPDFSrc, $GUI_ENABLE)
					Else
						GUICtrlSetState($hPDFSrc, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hPDFSrc
					Switch GUICtrlRead($hPDFSrc)
						Case "Default"
							If Not RunPDFCheck($bSilent) Then SetError(1)
							$sHandler = "Default"
						Case "Custom"
							$sHandler = FileOpenDialog("Select a PDF Handler", @ProgramFilesDir, "Executables (*.exe)", $FD_FILEMUSTEXIST)
					EndSwitch
					If @error Then
						GUICtrlSetState($hNoPDFs, $GUI_UNCHECKED)
						GUICtrlSetState($hPDFSrc, $GUI_DISABLE)
					EndIf

				Case Else
					;;;

			EndSelect

		WEnd

	EndIf

EndFunc

Func RunUpdateCheck($bFull = False)
	Switch _GetLatestRelease($sVersion)
		#cs
		Case -1
			MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
				_Translate($aMUI[1], "Test Build?"), _
				_Translate($aMUI[1], "You're running a newer build than publicly Available!"), _
				10)
		#ce
		Case 0
			If $bFull Then
				Switch @error
					Case 0
						MsgBox($MB_OK + $MB_ICONINFORMATION + $MB_TOPMOST, _
							_Translate($aMUI[1], "Up to Date"), _
							_Translate($aMUI[1], "You're running the latest build!"), _
							10)
					Case 1
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "Unable to Check for Updates"), _
							_Translate($aMUI[1], "Unable to load release data."), _
							10)
					Case 2
						MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
							_Translate($aMUI[1], "Unable to Check for Updates"), _
							_Translate($aMUI[1], "Invalid Data Received!"), _
							10)
					Case 3
						Switch @extended
							Case 0
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Unable to Check for Updates"), _
									_Translate($aMUI[1], "Invalid Release Tags Received!"), _
									10)
							Case 1
								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
									_Translate($aMUI[1], "Unable to Check for Updates"), _
									_Translate($aMUI[1], "Invalid Release Types Received!"), _
									10)
						EndSwitch
				EndSwitch
			EndIf
		Case 1
			If MsgBox($MB_YESNO + $MB_ICONINFORMATION + $MB_TOPMOST, _
				_Translate($aMUI[1], "MSEdgeRedirect Update Available"), _
				_Translate($aMUI[1], "An Update is Available, would you like to download it?"), _
				10) = $IDYES Then ShellExecute("https://fcofix.org/MSEdgeRedirect/releases/latest")
	EndSwitch
EndFunc

Func SetAppRegistry(ByRef $aConfig)

	Local Enum $vMode = 2

	Local $sHive = ""
	Local $sLocation = ""

	If $aConfig[$vMode] Then
		$sLocation = $sDrive & "\Program Files\MSEdgeRedirect\"
		$sHive = "HKLM"
	Else
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
		$sHive = "HKCU"
	EndIf

	; App Paths
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "", "REG_SZ", $sLocation & "MSEdgeRedirect.exe")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "Path", "REG_SZ", $sLocation)

	; URI Handler for Pre Win11 22494 Installs
	#cs
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSEdgeRedirect.exe", "SupportedProtocols", "REG_SZ", "microsoft-edge")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge", "", "REG_SZ", "URL:microsoft-edge")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge", "URL Protocol", "REG_SZ", "")
	RegWrite($sHive & "\Software\Classes\MSEdgeRedirect.microsoft-edge\shell\open\command", "", "REG_SZ", '"' & $sLocation & 'MSEdgeRedirect.exe" "%1"')
	#ce

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
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "URLUpdateInfo", "REG_SZ", "https://msedgeredirect.com/releases/latest")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Version", "REG_SZ", $sVersion)

EndFunc

Func SetAppShortcuts(ByRef $aConfig, ByRef $aSettings)

	Local $sArgs = ""
	Local Enum $vMode = 2

	If $aSettings[$bNoTray] Then $sArgs = "/hide"

	Switch $aSettings[$sStartMenu]

		Case "Full"
			If $aConfig[$vMode] Then
				DirCreate(@ProgramsCommonDir & "\MSEdgeRedirect")
				FileCreateShortcut($sDrive & "\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe", @ProgramsCommonDir & "\MSEdgeRedirect\MSER Settings.lnk", $sDrive & "\Program Files\MSEdgeRedirect\", "/settings")
			Else
				DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
				FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\MSER Settings.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", "/settings")
				ContinueCase
			EndIf

		Case "App Only"
			DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
			FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\MSEdgeRedirect.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", $sArgs)

		Case Else
			FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Unexpected StartMenu Value: " & $aSettings[$sStartMenu] & @CRLF)

	EndSwitch

EndFunc

Func SetIFEORegistry(ByRef $aChannels)

	RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe")
	RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe", "UseFilter", "REG_DWORD", 1)
	For $iLoop = 1 To $aEdges[0] Step 1
		If $aChannels[$iLoop - 1] Then
			RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop)
			RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop, "Debugger", "REG_SZ", $sDrive & "\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe")
			RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\MSER" & $iLoop, "FilterFullPath", "REG_SZ", $aEdges[$iLoop])	
			If $iLoop = $aEdges[0] Then
				;;;
			Else
				FileCreateNTFSLink(StringReplace($aEdges[$iLoop], "\msedge.exe", ""), StringReplace($aEdges[$iLoop], "Application\msedge.exe", "IFEO\"), $FC_OVERWRITE)
			EndIf
		EndIf
	Next
	If $aChannels[4] Then ; IE_TO_EDGE_STUB
		RegDelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msedge.exe\0")
		RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe")
		RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe", "UseFilter", "REG_DWORD", 1)
		RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe\0")
		RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe\0", "Debugger", "REG_SZ", $sDrive & "\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe")
		RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\ie_to_edge_stub.exe\0", "FilterFullPath", "REG_SZ", $aEdges[5])
	EndIf
EndFunc

Func SetOptionsRegistry($sName, $vValue, ByRef $aConfig)

	Local Static $sHive = ""
	Local Static $sPolicy = ""
	Local Enum $bManaged = 1, $vMode

	If $sHive = "" Then
		If $aConfig[$vMode] Then
			$sHive = "HKLM"
		Else
			$sHive = "HKCU"
		EndIf

		If $aConfig[$bManaged] Then $sPolicy = "Policies\"
	EndIf

	Select
		Case IsBool($vValue)
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_DWORD", $vValue)
			If @error Then FileWrite($hLogs[$AppFailures], _NowCalc() & " - [WARNING!] Unable to write REG_DWORD Registry Key '" & $sName & "' - with value '" & $vValue & "'" & @CRLF)

		Case IsString($vValue)
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_SZ", $vValue)
			If @error Then FileWrite($hLogs[$AppFailures], _NowCalc() & " - [WARNING!] Unable to write REG_SZ Registry Key '" & $sName & "' - with value '" & $vValue & "'" & @CRLF)

		Case Else
			RegWrite($sHive & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sName, "REG_SZ", $vValue)
			If @error Then FileWrite($hLogs[$AppFailures], _NowCalc() & " - [WARNING!] Unable to write REG_SZ Registry Key '" & $sName & "' - with value '" & $vValue & "'" & @CRLF)

	EndSelect

EndFunc

Func SetScheduledTask($aChannels) ; Deprecated

	Local $hTS
	Local $hTO
	Local Enum $bManaged = 1, $vMode

	Local $aTasks[5] = [4, _
		"Update Edge.xml", _
		"Update Edge Beta.xml", _
		"Update Edge Dev.xml", _
		"Update Edge Canary.xml"]

	DirCreate($sDrive & "\Program Files\MSEdgeRedirect\Assets")
	FileInstall(".\Assets\Task Scheduler Tasks\Update Edge.xml", $sDrive & "\Program Files\MSEdgeRedirect\Assets\Update Edge.xml" , $FC_OVERWRITE)
	FileInstall(".\Assets\Task Scheduler Tasks\Update Edge Beta.xml", $sDrive & "\Program Files\MSEdgeRedirect\Assets\Update Edge Beta.xml" ,$FC_OVERWRITE)
	FileInstall(".\Assets\Task Scheduler Tasks\Update Edge Canary.xml", $sDrive & "\Program Files\MSEdgeRedirect\Assets\Update Edge Canary.xml" ,$FC_OVERWRITE)
	FileInstall(".\Assets\Task Scheduler Tasks\Update Edge Dev.xml", $sDrive & "\Program Files\MSEdgeRedirect\Assets\Update Edge Dev.xml" ,$FC_OVERWRITE)

	$hTS = _TS_Open()
	_TS_FolderCreate($hTS, "\MSEdgeRedirect")
	For $iLoop = 1 To $aTasks[0] Step 1
		If $aChannels[$iLoop - 1] Then
			$hTO = _TS_TaskImportXML($hTS, 1, $sDrive & "\Program Files\MSEdgeRedirect\Assets\" & $aTasks[$iLoop])
			_TS_TaskRegister($hTS, "\MSEdgeRedirect", $aTasks[$iLoop], $hTO)
		EndIf
	Next
	_TS_Close($hTS)

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

Func _IsInstalled()

	Local $aReturn[3] = [False, "", ""]

	$aReturn[2] = RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayVersion")
	If @error Then
		$aReturn[2] = RegRead("HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "DisplayVersion")
		If @error Then
			;;;
		Else
			$aReturn[0] = True
			$aReturn[1] = "HKCU"
		EndIf
	Else
		$aReturn[0] = True
		$aReturn[1] = "HKLM"
	EndIf

	Return $aReturn

EndFunc

Func _IsPriviledgedInstall()

	Local $hTestFile

	If @ScriptDir = $sDrive & "\Program Files\MSEdgeRedirect" Then
		Return True
	ElseIf @LocalAppDataDir & "\MSEdgeRedirect" Then
		Return False
	Else
		$hTestFile = FileOpen(".\writetest", $FO_CREATEPATH)
		If @error Then
			Return True
		Else
			FileDelete($hTestFile)
			Return False
		EndIf
	EndIf
EndFunc
