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

Func _GetSettingValue($sSetting)

	Local $vReturn = Null
	Local Static $bPortable

	If $sSetting = "SetPortable" Then $bPortable = True

	Select

		Case RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD Or $REG_QWORD
					$vReturn = Number(RegRead("HKLM\SOFTWARE\Policies\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD Or $REG_QWORD
					$vReturn = Number(RegRead("HKLM\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD Or $REG_QWORD
					$vReturn = Number(RegRead("HKCU\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case Not IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, Null) = Null And Not $bPortable
			$vReturn = _Bool(IniRead(@LocalAppDataDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, False))

		Case Not IniRead(@ScriptDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, Null) = Null
			$vReturn = _Bool(IniRead(@ScriptDir & "\MSEdgeRedirect\Settings.ini", "Settings", $sSetting, False))

		Case Else
			$vReturn = False

	EndSelect

	Return $vReturn

EndFunc