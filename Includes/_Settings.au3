#include-once

#include <Date.au3>
#include <AutoItConstants.au3>

#include ".\_Logging.au3"

Global $bDebug = False
Global $bIsAdmin = IsAdmin()
Global $bIsWOW64 = _WinAPI_IsWow64Process()
Global $bIs64Bit = @AutoItX64

Global $sDrive = StringLeft(@WindowsDir, 2)

If $bIs64Bit Then
	Global $aEdges[6] = [5, _
		$sDrive & "\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
		$sDrive & "\Program Files (x86)\Microsoft\Edge Beta\Application\msedge.exe", _
		$sDrive & "\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe", _
		@LocalAppDataDir & "\Microsoft\Edge SXS\Application\msedge.exe"]
Else
	Global $aEdges[6] = [5, _
		$sDrive & "\Program Files\Microsoft\Edge\Application\msedge.exe", _
		$sDrive & "\Program Files\Microsoft\Edge Beta\Application\msedge.exe", _
		$sDrive & "\Program Files\Microsoft\Edge Dev\Application\msedge.exe", _
		@LocalAppDataDir & "\Microsoft\Edge SXS\Application\msedge.exe"]
EndIf

Select
	Case FileExists($sDrive & "\Scripts\ie_to_edge_stub.exe")
		$aEdges[5] = $sDrive & "\Scripts\ie_to_edge_stub.exe"
	Case FileExists($sDrive & "\ProgramData\ie_to_edge_stub.exe")
		$aEdges[5] = $sDrive & "\ProgramData\ie_to_edge_stub.exe"
	Case FileExists($sDrive & "\Users\Public\ie_to_edge_stub.exe")
		$aEdges[5] = $sDrive & "\Users\Public\ie_to_edge_stub.exe"
	Case Else
		$aEdges[5] = $sDrive & "\Scripts\ie_to_edge_stub.exe"
EndSelect

Func _Bool($sString)
	Switch $sString
		Case "True", 1
			Return True
		Case "False", 0
			Return False
		Case Else
			Return $sString
	EndSwitch
EndFunc

Func _GetSettingValue($sSetting, $sLocation = Null)

	Local $vReturn = Null
	Local Static $bPortable

	Switch $sSetting
		Case "IsPortable"
			Return $bPortable
		Case "PerUser"
			$vReturn = Number(RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", "PerUser"))
		Case "RunUnsafe"
			$vReturn = Number(RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", "RunUnsafe"))
		Case "SetPortable"
			$bPortable = True
		Case Else
			Select

				Case $sLocation = "Policy"
					ContinueCase

				Case RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
					Switch @extended
						Case $REG_SZ Or $REG_EXPAND_SZ
							$vReturn = RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
						Case $REG_DWORD Or $REG_QWORD
							$vReturn = Number(RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting))
						Case Else
							FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
					EndSwitch

				Case $sLocation = "HKLM"
					ContinueCase

				Case RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
					Switch @extended
						Case $REG_SZ Or $REG_EXPAND_SZ
							$vReturn = RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
						Case $REG_DWORD Or $REG_QWORD
							$vReturn = Number(RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
						Case Else
							FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
					EndSwitch

				Case $sLocation = "HKCU"
					ContinueCase

				Case RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
					Switch @extended
						Case $REG_SZ Or $REG_EXPAND_SZ
							$vReturn = RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
						Case $REG_DWORD Or $REG_QWORD
							$vReturn = Number(RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
						Case Else
							FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
					EndSwitch

				Case $sLocation = "Appdata"
					ContinueCase

				Case Not IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, Null) = Null And Not $bPortable
					$vReturn = _Bool(IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, False))

				Case $sLocation = "Portable"
					ContinueCase

				Case Not IniRead(@ScriptDir & "\Settings.ini", "Settings", $sSetting, Null) = Null
					$vReturn = _Bool(IniRead(@ScriptDir & "\Settings.ini", "Settings", $sSetting, False))

				Case Else
					$vReturn = False

			EndSelect
	EndSwitch

	Return $vReturn

EndFunc

Func _SetSettingsValue($sSetting, $vValue, $sLocation)

	Local $sPolicy = ""
	
	Switch $sLocation

		Case "Policy"
			$sPolicy = "Policies\"
			ContinueCase

		Case "HKLM", "HKCU"
			Select
				Case IsBool($vValue)
					RegWrite($sLocation & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sSetting, "REG_DWORD", $vValue)
					If @error Then FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unable to write " & $sLocation & " REG_DWORD Registry Key '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)
		
				Case IsString($vValue)
					RegWrite($sLocation & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sSetting, "REG_SZ", $vValue)
					If @error Then FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unable to write " & $sLocation & " REG_SZ Registry Key '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)
		
				Case Else
					RegWrite($sLocation & "\SOFTWARE\" & $sPolicy & "Robert Maehl Software\MSEdgeRedirect\", $sSetting, "REG_SZ", $vValue)
					If @error Then FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unable to write " & $sLocation & " REG_SZ Registry Key '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)
		
			EndSelect

		Case "Appdata"
			IniWrite(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, $vValue)
			FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unable to write '" & $sLocation & "' INI Key '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)

		Case "Portable"
			IniWrite(@LocalAppDataDir & "\Settings.ini", "Settings", $sSetting, $vValue)
			FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unable to write '" & $sLocation & "' INI Key '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)

		Case Else
			FileWrite($hLogs[$Install], _NowCalc() & " - [WARNING!] Unknown Settings Location '" & $sLocation & "' when attempting to write '" & $sSetting & "' - with value '" & $vValue & "'" & @CRLF)

	EndSwitch

EndFunc