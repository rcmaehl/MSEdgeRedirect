#include-once

#include <Date.au3>
#include <AutoItConstants.au3>

#include ".\_Logging.au3"

Global $bIsAdmin = IsAdmin()
Global $bIs64Bit = Not _WinAPI_IsWow64Process()

If $bIs64Bit Then
	Global $aEdges[5] = [4, _
		"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
		"C:\Program Files (x86)\Microsoft\Edge Beta\Application\msedge.exe", _
		"C:\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe", _
		@LocalAppDataDir & "\Microsoft\Edge SXS\Application\msedge.exe"]
Else
	Global $aEdges[5] = [4, _
		"C:\Program Files\Microsoft\Edge\Application\msedge.exe", _
		"C:\Program Files\Microsoft\Edge Beta\Application\msedge.exe", _
		"C:\Program Files\Microsoft\Edge Dev\Application\msedge.exe", _
		@LocalAppDataDir & "\Microsoft\Edge SXS\Application\msedge.exe"]
EndIf

Func _Bool($sString)
	If $sString = "True" Then
		Return True
	ElseIf $sString = "False" Then
		Return False
	Else
		Return $sString
	EndIf
EndFunc

Func _GetSettingValue($sSetting, $bPortable = False)

	Local $vReturn = Null

	Local $sHive1 = ""
	Local $sHive2 = ""

	If $bIs64Bit Then
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
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD
					$vReturn = Number(RegRead($sHive1 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
			EndSwitch

		Case RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting) And Not $bPortable
			Switch @extended
				Case $REG_SZ Or $REG_EXPAND_SZ
					$vReturn = RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting)
				Case $REG_DWORD
					$vReturn = Number(RegRead($sHive2 & "\SOFTWARE\Robert Maehl Software\MSEdgeRedirect", $sSetting))
				Case Else
					FileWrite($hLogs[$AppFailures], _NowCalc() & " - Invalid Registry Key Type: " & $sSetting & @CRLF)
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