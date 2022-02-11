#include-once

#include <Misc.au3>
#include <String.au3>
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

Global $sVersion

If @Compiled Then
	$sVersion = FileGetVersion(@ScriptFullPath)
Else
	$sVersion = "x.x.x.x"
EndIf

Func RunInstall(ByRef $aConfig, ByRef $aSettings)

	Local $sArgs = ""
	Local Enum $bManaged = 1, $vMode
	Local Enum $bNoApps, $bNoBing, $bNoMSN, $bNoPDFs, $bNoTray, $bNoUpdates, $sPDFApp, $sSearch, $sSearchPath, $bStartup = 10, $sWeather

	SetOptionsRegistry("NoApps", $aSettings[$bNoApps], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoBing", $aSettings[$bNoBing], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoMSN", $aSettings[$bNoMSN], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoPDFs", $aSettings[$bNoPDFs], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoTray", $aSettings[$bNoTray], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("NoUpdates", $aSettings[$bNoUpdates], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("PDFApp", $aSettings[$sPDFApp], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("Search", $aSettings[$sSearch], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("SearchPath", $aSettings[$sSearchPath], $aConfig[$vMode], $aConfig[$bManaged])
	SetOptionsRegistry("Weather", $aSettings[$sWeather], $aConfig[$vMode], $aConfig[$bManaged])

	If $aConfig[$vMode] Then
		FileCopy(@ScriptFullPath, "C:\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE)
	Else
		If $aSettings[$bNoTray] Then $sArgs = "/hide"
		FileCopy(@ScriptFullPath, @LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", $FC_CREATEPATH+$FC_OVERWRITE)
		If $aSettings[$bStartup] Then FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @StartupDir & "\MSEdgeRedirect.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", $sArgs)
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

	If $bIsAdmin Then
		$sLocation = "C:\Program Files\MSEdgeRedirect\"
		If $bIs64Bit Then
			$sHive = "HKLM64"
		Else
			$sHive = "HKLM"
		EndIf
	Else
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
		If $bIs64Bit Then
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

	; Start Menu Shortcuts
	FileDelete(@StartupDir & "\MSEdgeRedirect.lnk")
	DirRemove(@ProgramsCommonDir & "\MSEdgeRedirect", $DIR_REMOVE)
	DirRemove(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect", $DIR_REMOVE)

	If $bIsAdmin Then
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

	If $bIsAdmin Then
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

Func RunSetup($bUpdate = False, $bSilent = False, $iPage = 0, $hSetupFile = @ScriptDir & "\Setup.ini")

	Local $hMsg
	Local $sArgs = ""
	Local $iMode = $iPage
	Local $sEdges
	Local $sEngine
	Local $sHandler
	Local $hChannels[4]
	Local $aChannels[4] = [True, False, False, False]

	Local $aConfig[3] = [$hSetupFile, False, "Service"] ; Default Setup.ini Values
	Local Enum $hFile, $bManaged, $vMode

	Local $aSettings[12] = [False, False, False, False, False, False, "", "", "", "Full", True, ""]
	Local Enum $bNoApps, $bNoBing, $bNoMSN, $bNoPDFs, $bNoTray, $bNoUpdates, $sPDFApp, $sSearch, $sSearchPath, $sStartMenu, $bStartup, $sWeather

	If $bSilent Then

		If $aConfig[$hFile] = "WINGET" Then
			; Bypass file checks, IniReads, use default values
		ElseIf Not FileExists($aConfig[$hFile]) Then
			Exit 2 ; ERROR_FILE_NOT_FOUND
		Else
			$aConfig[$bManaged] = _Bool(IniRead($aConfig[$hFile], "Config", "Managed", False))
			$aConfig[$vMode] = IniRead($aConfig[$hFile], "Config", "Mode", "Service")

			If $aConfig[$vMode] = "active" Then
				$aConfig[$vMode] = True
			Else
				$aConfig[$vMode] = False
			EndIf

			$aSettings[$bNoApps] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoApps", $aSettings[$bNoApps]))
			$aSettings[$bNoBing] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoBing", $aSettings[$bNoBing]))
			$aSettings[$bNoMSN] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoMSN", $aSettings[$bNoBing]))
			$aSettings[$bNoPDFs] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoPDFs", $aSettings[$bNoPDFs]))
			$aSettings[$bNoTray] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoTray", $aSettings[$bNoTray]))
			$aSettings[$bNoUpdates] = _Bool(IniRead($aConfig[$hFile], "Settings", "NoUpdates", $aSettings[$bNoUpdates]))
			$aSettings[$sPDFApp] = IniRead($aConfig[$hFile], "Settings", "PDFApp", $aSettings[$sPDFApp])
			$aSettings[$sSearch] = IniRead($aConfig[$hFile], "Settings", "Search", $aSettings[$sSearch])
			$aSettings[$sSearchPath] = IniRead($aConfig[$hFile], "Settings", "SearchPath", $aSettings[$sSearchPath])
			$aSettings[$sStartMenu] = IniRead($aConfig[$hFile], "Settings", "StartMenu", $aSettings[$sStartMenu])
			$aSettings[$bStartup] = _Bool(IniRead($aConfig[$hFile], "Settings", "Startup", $aSettings[$bStartup]))
			$aSettings[$sWeather] = IniRead($aConfig[$hFile], "Settings", "Weather", $aSettings[$sSearch])
		EndIf

		If ($aConfig[$bManaged] Or $aConfig[$vMode]) And Not $bIsAdmin Then Exit 5 ; ERROR_ACCESS_DENIED

		$sEdges = IniRead($aConfig[$hFile], "Settings", "Edges", "")
		If StringInStr($sEdges, "Stable") Then $aChannels[0] = True
		If StringInStr($sEdges, "Beta") Then $aChannels[1] = True
		If StringInStr($sEdges, "Dev") Then $aChannels[2] = True
		If StringInStr($sEdges, "Canary") Then $aChannels[3] = True

		For $iLoop = 0 To 3 Step 1
			If $aChannels[$iLoop] = True Then ExitLoop
			If $iLoop = 3 Then Exit 160 ; ERROR_BAD_ARGUMENTS
		Next

		For $iLoop = $bNoApps To $bNoUpdates Step 1
			If Not IsBool($aSettings[$iLoop]) Then Exit 160 ; ERROR_BAD_ARGUMENTS
		Next
		If Not IsBool($aSettings[$bStartup]) Then Exit 160 ; ERROR_BAD_ARGUMENTS

		If $bUpdate Then RunRemoval(True)
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

		Local $hInstallGUI = GUICreate("MSEdge Redirect " & $sVersion & " Setup", 640, 480)

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

		GUISwitch($hInstallGUI)
		#EndRegion

		#Region Mode Page
		$aPages[$hMode] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)
		If $bUpdate Then
			GUICtrlCreateLabel("MSEdge Redirect " & $sVersion & " Update", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Next to continue the Update of MS Edge Redirect after customizing your preferred mode", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Install MSEdge Redirect " & $sVersion, 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Next to continue the Install of MS Edge Redirect after customizing your preferred mode", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		GUICtrlCreateGroup("Mode", 20, 80, 420, 320)
			Local $hService = GUICtrlCreateRadio("Service Mode - Per User" & @CRLF & _
				@CRLF & _
				"MSEdge Redirect stays running in the background. Detected Edge data is redirected to your default browser.", _
				50, 100, 380, 60, $BS_TOP+$BS_MULTILINE)
			If Not $bIsAdmin Then GUICtrlSetState(-1, $GUI_CHECKED)

			Local $hStartup = GUICtrlCreateCheckbox("Start MSEdge Redirect Service With Windows", 70, 160, 320, 20)
			Local $hNoIcon = GUICtrlCreateCheckbox("Hide MSEdge Redirect Service Icon from Tray", 70, 180, 320, 20)

			If $bIsAdmin Then
				GUICtrlSetState($hStartup, $GUI_DISABLE)
				GUICtrlSetState($hNoIcon, $GUI_DISABLE)
			EndIf

			GUICtrlCreateIcon("imageres.dll", 78, 30, 210, 16, 16)
			Local $hActive = GUICtrlCreateRadio("Active Mode - All Users" & @CRLF & _
				@CRLF & _
				"MSEdge Redirect only runs when a selected Edge is launched, similary to the old EdgeDeflector app.", _
				50, 210, 380, 60, $BS_TOP+$BS_MULTILINE)
			If $bIsAdmin Then GUICtrlSetState(-1, $GUI_CHECKED)

			$hChannels[0] = GUICtrlCreateCheckbox("Edge Stable", 70, 270, 90, 20)
			GUICtrlSetState(-1, $GUI_CHECKED)
			$hChannels[1] = GUICtrlCreateCheckbox("Edge Beta", 160, 270, 90, 20)
			$hChannels[2] = GUICtrlCreateCheckbox("Edge Dev", 250, 270, 90, 20)
			$hChannels[3] = GUICtrlCreateCheckbox("Edge Canary", 340, 270, 90, 20)

			If Not $bIsAdmin Then
				GUICtrlSetState($hChannels[0], $GUI_DISABLE)
				GUICtrlSetState($hChannels[1], $GUI_DISABLE)
				GUICtrlSetState($hChannels[2], $GUI_DISABLE)
				GUICtrlSetState($hChannels[3], $GUI_DISABLE)
			EndIf
		#EndRegion

		#Region Settings Page
		$aPages[$hSettings] = GUICreate("", 460, 420, 180, 0, $WS_POPUP, $WS_EX_MDICHILD, $hInstallGUI)
		GUISetBkColor(0xFFFFFF)

		If $bUpdate Then
			GUICtrlCreateLabel("MSEdge Redirect " & $sVersion & " Update", 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Install to continue the Update of MS Edge Redirect after customizing your preferred settings", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		Else
			GUICtrlCreateLabel("Install MSEdge Redirect " & $sVersion, 20, 10, 420, 30)
			GUICtrlSetFont(-1, 20, $FW_BOLD, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
			GUICtrlCreateLabel("Click Install to continue the Install of MS Edge Redirect after customizing your preferred settings", 20, 40, 420, 40)
			GUICtrlSetFont(-1, 10, $FW_NORMAL, $GUI_FONTNORMAL, "", $CLEARTYPE_QUALITY)
		EndIf

		GUICtrlCreateGroup("Options", 20, 80, 420, 320)
			Local $hNoApps = GUICtrlCreateCheckbox("De-embed Windows Store 'Apps'", 50, 100, 380, 20)
			Local $hNoPDFs = GUICtrlCreateCheckbox("Redirect PDFs to:", 50, 120, 380, 20)
			Local $hPDFPath = GUICtrlCreateEdit("", 50, 140, 380, 20, $ES_READONLY)
			Local $hSearch = GUICtrlCreateCheckbox("Replace Bing Search Results with:", 50, 165, 380, 20)
			Local $hEngine = GUICtrlCreateCombo("", 50, 185, 380, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "Ask|Baidu|Custom|DuckDuckGo|Ecosia|Google|Sogou|Yahoo|Yandex", "Google")
			GUICtrlSetState(-1, $GUI_DISABLE)
			Local $hNoMSN = GUICtrlCreateCheckbox("Replace Weather Results with:", 50, 210, 380, 20)
			Local $hWeather = GUICtrlCreateCombo("", 50, 230, 380, 20, $CBS_DROPDOWNLIST+$WS_VSCROLL)
			GUICtrlSetData(-1, "AccuWeather|Weather.com|Weather.gov|Windy|WUnderground|Ventusky", "Weather.com")
			GUICtrlSetState(-1, $GUI_DISABLE)

		If $bUpdate Then
			GUICtrlSetState($hNoApps, _GetSettingValue("NoApps"))
			GUICtrlSetState($hNoPDFs, _GetSettingValue("NoPDFs"))
			If _IsChecked($hNoPDFs) Then
				$sHandler = _GetSettingValue("PDFApp")
				GUICtrlSetData($hPDFPath, " " & $sHandler)
			EndIf
			GUICtrlSetState($hSearch, _GetSettingValue("NoBing"))
			If _IsChecked($hSearch) Then
				GUICtrlSetState($hEngine, $GUI_ENABLE)
				GUICtrlSetData($hEngine, _GetSettingValue("Search"))
				$sEngine = _GetSettingValue("SearchPath")
			EndIf
			GUICtrlSetState($hNoMSN, _GetSettingValue("NoMSN"))
			If _IsChecked($hNoMSN) Then
				GUICtrlSetState($hWeather, $GUI_ENABLE)
				GUICtrlSetData($hWeather, _GetSettingValue("Weather"))
			EndIf
		EndIf

		GUISwitch($hInstallGUI)

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
					If $iMode = $hSettings Then Exit
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
							If $bUpdate Then RunRemoval(True)

							$aConfig[$vMode] = _IsChecked($hActive)

							$aSettings[$bNoApps] = _IsChecked($hNoApps)
							$aSettings[$bNoBing] = _IsChecked($hSearch)
							$aSettings[$bNoMSN] = _IsChecked($hNoMSN)
							$aSettings[$bNoPDFs] = _IsChecked($hNoPDFs)
							$aSettings[$bNoTray] = _IsChecked($hNoIcon)
							$aSettings[$sPDFApp] = $sHandler
							$aSettings[$sSearch] = GUICtrlRead($hEngine)
							$aSettings[$sSearchPath] = $sEngine
							$aSettings[$bStartup] = _IsChecked($hStartup)
							$aSettings[$sWeather] = GUICtrlRead($hWeather)

							GUISetState(@SW_HIDE, $hSettings)
							RunInstall($aConfig, $aSettings)
							SetAppRegistry($aConfig[$vMode])
							If $aConfig[$vMode] Then
								For $iLoop = 0 To 3 Step 1
									$aChannels[$iLoop] = _IsChecked($hChannels[$iLoop])
								Next
								SetIFEORegistry($aChannels)
							EndIf
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
						GUICtrlSetState($hChannels[0], $GUI_ENABLE)
						GUICtrlSetState($hChannels[1], $GUI_ENABLE)
						GUICtrlSetState($hChannels[2], $GUI_ENABLE)
						GUICtrlSetState($hChannels[3], $GUI_ENABLE)
						ContinueCase
					EndIf

				Case $hMsg = $hChannels[0] Or $hMsg = $hChannels[1] Or $hMsg = $hChannels[2] Or $hMsg = $hChannels[3]
					;GUICtrlSetState($hInstall, $GUI_DISABLE)
					For $iLoop = 0 To 3 Step 1
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
					$sEngine = InputBox("Enter Search Engine URL", "Enter the URL format of the custom search Engine to use", "https://duckduckgo.com/?q=")
					If @error Then GUICtrlSetData($hEngine, "Google")

				Case $hMsg = $hNoMSN
					If _IsChecked($hNoMSN) Then
						GUICtrlSetState($hWeather, $GUI_ENABLE)
					Else
						GUICtrlSetState($hWeather, $GUI_DISABLE)
					EndIf

				Case $hMsg = $hNoPDFs
					If _IsChecked($hNoPDFs) Then
						$sHandler = FileOpenDialog("Select a PDF Handler", @ProgramFilesDir, "Executables (*.exe)", $FD_FILEMUSTEXIST)
						If @error Then
							GUICtrlSetState($hNoPDFs, $GUI_UNCHECKED)
						Else
							GUICtrlSetData($hPDFPath, " " & $sHandler)
						EndIf
					Else
						GUICtrlSetData($hPDFPath, "")
					EndIf

				Case Else
					;;;

			EndSelect

		WEnd

	EndIf

EndFunc

Func RunUpdateCheck($bFull = False)
	Switch _GetLatestRelease($sVersion)
		Case -1
			MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, _
				_Translate($aMUI[1], "Test Build?"), _
				_Translate($aMUI[1], "You're running a newer build than publicly Available!"), _
				10)
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
				_Translate($aMUI[1], "Update Available"), _
				_Translate($aMUI[1], "An Update is Available, would you like to download it?"), _
				10) = $IDYES Then ShellExecute("https://fcofix.org/MSEdgeRedirect/releases")
	EndSwitch
EndFunc

Func SetAppRegistry($bAllUsers)

	Local $sHive = ""
	Local $sLocation = ""

	If $bAllUsers Then
		$sLocation = "C:\Program Files\MSEdgeRedirect\"
		If $bIs64Bit Then
			$sHive = "HKLM64"
		Else
			$sHive = "HKLM"
		EndIf
	Else
		$sLocation = @LocalAppDataDir & "\MSEdgeRedirect\"
		If $bIs64Bit Then
			$sHive = "HKCU64"
		Else
			$sHive = "HKCU"
		EndIf
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
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "URLUpdateInfo", "REG_SZ", "https://msedgeredirect.com/releases")
	RegWrite($sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSEdgeRedirect", "Version", "REG_SZ", $sVersion)

EndFunc

Func SetAppShortcuts(ByRef $aConfig, ByRef $aSettings)

	Local $sArgs = ""
	Local Enum $vMode = 2
	Local Enum $bNoTray = 4, $sStartMenu = 9

	If $aSettings[$bNoTray] Then $sArgs = "/hide"

	Switch $aSettings[$sStartMenu]

		Case "Full"
			If $aConfig[$vMode] Then
				DirCreate(@ProgramsCommonDir & "\MSEdgeRedirect")
				FileCreateShortcut("C:\Program Files\MSEdgeRedirect\MSEdgeRedirect.exe", @ProgramsCommonDir & "\MSEdgeRedirect\Settings.lnk", "C:\Program Files\MSEdgeRedirect\", "/settings")
			Else
				DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
				FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\Settings.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", "/settings")
				ContinueCase
			EndIf

		Case "App Only"
			DirCreate(@AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect")
			FileCreateShortcut(@LocalAppDataDir & "\MSEdgeRedirect\MSEdgeRedirect.exe", @AppDataDir & "\Microsoft\Windows\Start Menu\Programs\MSEdgeRedirect\MSEdgeRedirect.lnk", @LocalAppDataDir & "\MSEdgeRedirect\", $sArgs)

		Case Else
			;;;

	EndSwitch

EndFunc

Func SetIFEORegistry(ByRef $aChannels)

	Local $sHive = ""

	If $bIs64Bit Then
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
			If $bIs64Bit Then
				$sHive = "HKLM64"
			Else
				$sHive = "HKLM"
			EndIf
		Else
			If $bIs64Bit Then
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

	Local $sHive1 = ""
	Local $sHive2 = ""
	Local $aReturn[3] = [False, "", ""]

	If $bIs64Bit Then
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
