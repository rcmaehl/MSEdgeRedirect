#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Assets\MSEdgeRedirect.ico
#AutoIt3Wrapper_Outfile=MSEdgeRedirect_x86.exe
#AutoIt3Wrapper_Outfile_x64=MSEdgeRedirect.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=A Tool to Redirect New, Search, and Weather Results to Your Default Browser
#AutoIt3Wrapper_Res_Fileversion=0.1.0.0
#AutoIt3Wrapper_Res_ProductName=MSEdgeRedirect
#AutoIt3Wrapper_Res_ProductVersion=0.1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Robert Maehl, using LGPL 3 License
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 -v1 -v2 -v3
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>

Opt("TrayAutoPause", 0)

Local $aProcessList
Local $aAdjust, $aList = 0

; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))

_WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

While 1
	If ProcessExists("msedge.exe") Then
		$aProcessList = ProcessList("msedge.exe")
		For $iLoop = 1 To $aProcessList[0][0] - 1
			$sCommandline = _WinAPI_GetProcessCommandLine($aProcessList[$iLoop][1])
			If StringInStr($sCommandline, "--single-argument") And _WinAPI_GetProcessFileName($aProcessList[$iLoop][1]) = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" Then
				Select
					Case StringInStr($sCommandline, "Microsoft.Windows.Cortana")
						$sCommandline = StringReplace($sCommandline, "--single-argument microsoft-edge:?launchContext1=Microsoft.Windows.Cortana_cw5n1h2txyewy&url=", "")
						$sCommandline = _UnicodeURLDecode($sCommandline)
						If _WinAPI_UrlIs($sCommandline) Then ShellExecute($sCommandline)
					Case Else
						$sCommandline = StringReplace($sCommandline, "--single-argument microsoft-edge:", "")
						If _WinAPI_UrlIs($sCommandline) Then ShellExecute($sCommandline)
				EndSelect
				ProcessClose($aProcessList[$iLoop][1])
			EndIf
		Next
	EndIf
WEnd

_WinAPI_AdjustTokenPrivileges($hToken, $aAdjust, 0, $aAdjust)
_WinAPI_CloseHandle($hToken)

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
    $Process = StringToBinary (StringReplace($strChar, "+", " "))
    $DecodedString = BinaryToString ($Process, 4)
    Return $DecodedString
EndFunc   ;==>_UnicodeURLDecode