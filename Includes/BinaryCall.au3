; =============================================================================
;  AutoIt BinaryCall UDF (2014.7.24)
;  Purpose: Allocate, Decompress, And Prepare Binary Machine Code
;  Author: Ward
; =============================================================================

#Include-once

Global $__BinaryCall_Kernel32dll = DllOpen('kernel32.dll')
Global $__BinaryCall_Msvcrtdll = DllOpen('msvcrt.dll')
Global $__BinaryCall_LastError = ""

Func _BinaryCall_GetProcAddress($Module, $Proc)
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, 'ptr', 'GetProcAddress', 'ptr', $Module, 'str', $Proc)
	If @Error Or Not $Ret[0] Then Return SetError(1, @Error, 0)
	Return $Ret[0]
EndFunc

Func _BinaryCall_LoadLibrary($Filename)
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, "handle", "LoadLibraryW", "wstr", $Filename)
	If @Error Then Return SetError(1, @Error, 0)
	Return $Ret[0]
EndFunc

Func _BinaryCall_lstrlenA($Ptr)
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, "int", "lstrlenA", "ptr", $Ptr)
	If @Error Then Return SetError(1, @Error, 0)
	Return $Ret[0]
EndFunc

Func _BinaryCall_Alloc($Code, $Padding = 0)
	Local $Length = BinaryLen($Code) + $Padding
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, "ptr", "VirtualAlloc", "ptr", 0, "ulong_ptr", $Length, "dword", 0x1000, "dword", 0x40)
	If @Error Or Not $Ret[0] Then Return SetError(1, @Error, 0)
	If BinaryLen($Code) Then
		Local $Buffer = DllStructCreate("byte[" & $Length & "]", $Ret[0])
		DllStructSetData($Buffer, 1, $Code)
	EndIf
	Return $Ret[0]
EndFunc

Func _BinaryCall_RegionSize($Ptr)
	Local $Buffer = DllStructCreate("ptr;ptr;dword;uint_ptr;dword;dword;dword")
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, "int", "VirtualQuery", "ptr", $Ptr, "ptr", DllStructGetPtr($Buffer), "uint_ptr", DllStructGetSize($Buffer))
	If @Error Or $Ret[0] = 0 Then Return SetError(1, @Error, 0)
	Return DllStructGetData($Buffer, 4)
EndFunc

Func _BinaryCall_Free($Ptr)
	Local $Ret = DllCall($__BinaryCall_Kernel32dll, "bool", "VirtualFree", "ptr", $Ptr, "ulong_ptr", 0, "dword", 0x8000)
	If @Error Or $Ret[0] = 0 Then
		$Ret = DllCall($__BinaryCall_Kernel32dll, "bool", "GlobalFree", "ptr", $Ptr)
		If @Error Or $Ret[0] <> 0 Then Return SetError(1, @Error, False)
	EndIf
	Return True
EndFunc

Func _BinaryCall_Release($CodeBase)
	Local $Ret = _BinaryCall_Free($CodeBase)
	Return SetError(@Error, @Extended, $Ret)
EndFunc

Func _BinaryCall_MemorySearch($Ptr, $Length, $Binary)
	Static $CodeBase
	If Not $CodeBase Then
		If @AutoItX64 Then
			$CodeBase = _BinaryCall_Create('0x4883EC084D85C94889C8742C4C39CA72254C29CA488D141131C9EB0848FFC14C39C97414448A1408453A140874EE48FFC04839D076E231C05AC3', '', 0, True, False)
		Else
			$CodeBase = _BinaryCall_Create('0x5589E58B4D14578B4508568B550C538B7D1085C9742139CA721B29CA8D341031D2EB054239CA740F8A1C17381C1074F34039F076EA31C05B5E5F5DC3', '', 0, True, False)
		EndIf
		If Not $CodeBase Then Return SetError(1, 0, 0)
	EndIf

	$Binary = Binary($Binary)
	Local $Buffer = DllStructCreate("byte[" & BinaryLen($Binary) & "]")
	DllStructSetData($Buffer, 1, $Binary)

	Local $Ret = DllCallAddress("ptr:cdecl", $CodeBase, "ptr", $Ptr, "uint", $Length, "ptr", DllStructGetPtr($Buffer), "uint", DllStructGetSize($Buffer))
	Return $Ret[0]
EndFunc

Func _BinaryCall_Base64Decode($Src)
	Static $CodeBase
	If Not $CodeBase Then
		If @AutoItX64 Then
			$CodeBase = _BinaryCall_Create('0x41544989CAB9FF000000555756E8BE000000534881EC000100004889E7F3A44C89D6E98A0000004439C87E0731C0E98D0000000FB66E01440FB626FFC00FB65E020FB62C2C460FB62424408A3C1C0FB65E034189EB41C1E4024183E3308A1C1C41C1FB044509E34080FF634189CC45881C08744C440FB6DFC1E5044489DF4088E883E73CC1FF0209C7418D44240241887C08014883C10380FB63742488D841C1E3064883C60483E03F4409D841884408FF89F389C84429D339D30F8C67FFFFFF4881C4000100005B5E5F5D415CC35EC3E8F9FFFFFF000000000000000000000000000000000000000000000000000000000000000000000000000000000000003E0000003F3435363738393A3B3C3D00000063000000000102030405060708090A0B0C0D0E0F101112131415161718190000000000001A1B1C1D1E1F202122232425262728292A2B2C2D2E2F30313233', '', 132, True, False)
		Else
			$CodeBase = _BinaryCall_Create('0x55B9FF00000089E531C05756E8F10000005381EC0C0100008B55088DBDF5FEFFFFF3A4E9C00000003B45140F8FC20000000FB65C0A028A9C1DF5FEFFFF889DF3FEFFFF0FB65C0A038A9C1DF5FEFFFF889DF2FEFFFF0FB65C0A018985E8FEFFFF0FB69C1DF5FEFFFF899DECFEFFFF0FB63C0A89DE83E630C1FE040FB6BC3DF5FEFFFFC1E70209FE8B7D1089F3881C074080BDF3FEFFFF63745C0FB6B5F3FEFFFF8BBDECFEFFFF8B9DE8FEFFFF89F083E03CC1E704C1F80209F88B7D1088441F0189D883C00280BDF2FEFFFF6374278A85F2FEFFFFC1E60683C10483E03F09F088441F0289D883C0033B4D0C0F8C37FFFFFFEB0231C081C40C0100005B5E5F5DC35EC3E8F9FFFFFF000000000000000000000000000000000000000000000000000000000000000000000000000000000000003E0000003F3435363738393A3B3C3D00000063000000000102030405060708090A0B0C0D0E0F101112131415161718190000000000001A1B1C1D1E1F202122232425262728292A2B2C2D2E2F30313233', '', 132, True, False)
		EndIf
		If Not $CodeBase Then Return SetError(1, 0, Binary(""))
	EndIf

	$Src = String($Src)
	Local $SrcLen = StringLen($Src)
	Local $SrcBuf = DllStructCreate("char[" & $SrcLen & "]")
	DllStructSetData($SrcBuf, 1, $Src)

	Local $DstLen = Int(($SrcLen + 2) / 4) * 3 + 1
	Local $DstBuf = DllStructCreate("byte[" & $DstLen & "]")

	Local $Ret = DllCallAddress("uint:cdecl", $CodeBase, "ptr", DllStructGetPtr($SrcBuf), "uint", $SrcLen, "ptr", DllStructGetPtr($DstBuf), "uint", $DstLen)
	If $Ret[0] = 0 Then Return SetError(2, 0, Binary(""))
	Return BinaryMid(DllStructGetData($DstBuf, 1), 1, $Ret[0])
EndFunc

Func _BinaryCall_Base64Encode($Src)
	Static $CodeBase
	If Not $CodeBase Then
		If @AutoItX64 Then
			$CodeBase = _BinaryCall_Create('AwAAAARiAQAAAAAAAAArkuFQDAlvIp0qAgbDnjr76UDZs1EPNIP2K18t9s6SNTbd43IB7HfdyPM8VfD/o36z4AmSW2m2AIsC6Af3fKNsHU4BdQKGd0PQXHxPSX0iNqp1YAKovksqQna06NeKMoOYqryTUX4WgpHjokhp6zY2sEFSIjcL7dW3FDoNVz4bGPyZHRvjFwmqvr7YGlNYKwNoh+SYCXmIgVPVZ63Vz1fbT33/QFpWmWOeBRqs4J+c8Qp6zJFsK345Pjw0I8kMSsnho4F4oNzQ2OsAbmIioaQ6Ma2ziw5NH+M+t4SpEeHDnBdUTTL20sxWZ0yKruFAsBIRoHvM7LYcid2eBV2d5roSjnkwMG0g69LNjs1fHjbI/9iU/hJwpSsgl4fltXdZG659/li13UFY89M7UfckiZ9XOeBM0zadgNsy8r8M3rEAAA==')
		Else
			$CodeBase = _BinaryCall_Create('AwAAAARVAQAAAAAAAAAqr7blBndrIGnmhhfXD7R1fkOTKhicg1W6MCtStbz+CsneBEg0bbHH1sqTLmLfY7A6LqZl6LYWT5ULVj6MXgugPbBn9wKsSU2ZCcBBPNkx09HVPdUaKnbqghDGj/C5SHoF+A/5g+UgE1C5zJZORjJ8ljs5lt2Y9lA4BsY7jVKX2vmDvHK1NnSR6nVwh7Pb+Po/UpNcy5sObVWDKkYSCCtCIjKIYqOe3c6k8Xsp4eritCUprXEVvCFi7K5Z6HFXdm3nZsFcE+eSJ1WkRnVQbWcmpjGMGka61C68+CI7tsQ13UnCFWNSpDrCbzUejMZh8HdPgEc5vCg3pKMKin/NavNpB6+87Y9y7HIxmKsPdjDT30u9hUKWnYiRe3nrwKyVDsiYpKU/Nse368jHag5B5or3UKA+nb2+eY8JwzgA')
		EndIf
		If Not $CodeBase Then Return SetError(1, 0, Binary(""))
	EndIf

	$Src = Binary($Src)
	Local $SrcLen = BinaryLen($Src)
	Local $SrcBuf = DllStructCreate("byte[" & $SrcLen & "]")
	DllStructSetData($SrcBuf, 1, $Src)

	Local $DstLen = Int(($SrcLen + 2) / 3) * 4 + 1
	Local $DstBuf = DllStructCreate("char[" & $DstLen & "]")

	Local $Ret = DllCallAddress("uint:cdecl", $CodeBase, "ptr", DllStructGetPtr($SrcBuf), "uint", $SrcLen, "ptr", DllStructGetPtr($DstBuf), "uint", $DstLen)
	If $Ret[0] = 0 Then Return Binary("")
	Return StringMid(DllStructGetData($DstBuf, 1), 1, $Ret[0])
EndFunc

Func _BinaryCall_LzmaDecompress($Src)
	Static $CodeBase
	If Not $CodeBase Then
		If @AutoItX64 Then
			$CodeBase = _BinaryCall_Create(_BinaryCall_Base64Decode('QVcxwEFWQVVBVFVXSInXVkiJzlMx20iB7OgAAABEiiFBgPzgdgnpyQAAAEGD7C1BiMf/wEGA/Cx38THA6wRBg+wJQYjG/8BBgPwId/GLRglEi24FQQ+2zkyJRCQoRQ+2/0HB5xBBiQFBD7bEAcG4AAMAANPgjYQAcA4AAEhjyOjIBAAATInpSInF6L0EAABIicMxwEyJ8kSI4EyLRCQoiNQl//8A/0QJ+EiF24lFAHQoTYXtdCNIjVfzSI1MJDhIg8YNTIkEJE2J6UmJ2EiJ7+g2AAAAicbrBb4BAAAASInp6IQEAACF9nQKSInZMdvodgQAAEiJ2EiBxOgAAABbXl9dQVxBXUFeQV/DVVNBV0FWQVVBVEFQTQHBQVFNicVRVkgB8lJIieX8SYn0iwdMjX8Eik8Cg8r/0+L30olV6Ijhg8r/0+L30olV5ADBiUXsuAEAAACJReCJRdyJRdhIiUXQRSnJKfaDy/8A0bgAAwAA0+BIjYg2BwAAuAAEAARMif/R6fOrvwUAAADoUAMAAP/PdfdEie9EicgrfSDB4ARBifpEI1XoRAHQTY0cR+hAAwAAD4WTAAAAik3sI33k0+eA6Qj22dPuAfe4AQAAAEiNPH++AAEAAMHnCEGD+QdNjbR/bA4AAHI0TInvSCt90A+2P9HnQYnzIf5BAfNPjRxe6O8CAACJwcHuCIPhATnOvgABAAB1DjnGd9jrDE2J8+jQAgAAOfBy9EyJ76pEiclBg/kEcg65AwAAAEGD+QpyA4PBA0EpyelDAgAAT42cT4ABAADomgIAAHUsi0XcQYP5B4lF4BnAi1XY99CLTdCD4AOJVdxBicGJTdhNjbdkBgAA6akAAABPjZxPmAEAAOhfAgAAdUZEicjB4AREAdBNjZxH4AEAAOhHAgAAdWpBg/kHuQkAAAByA4PBAkGJyUyJ70grfdBIO30gD4L9AQAAigdIA33QqumzAQAAT42cT7ABAADoCgIAAIt12HQhT42cT8gBAADo+AEAAIt13HQJi03ci3XgiU3gi03YiU3ci03QiU3YiXXQQYP5B7kIAAAAcgODwQNBiclNjbdoCgAATYnz6LsBAAB1FESJ0CnJweADvggAAABJjXxGBOs2TY1eAuicAQAAdRpEidC5CAAAAMHgA74IAAAASY28RgQBAADrEUmNvgQCAAC5EAAAAL4AAQAAiU3MuAEAAABJifvoYQEAAInCKfJy8gNVzEGD+QSJVcwPg7kAAABBg8EHuQMAAAA50XICidHB4Qa4AQAAAEmNvE9gAwAAvkAAAABJifvoHwEAAEGJwkEp8nLwQYP6BHJ4RInWRIlV0NHug2XQAf/Og03QAkGD+g5zFYnx0mXQi0XQRCnQTY20R14FAADrLIPuBOi6AAAA0evRZdBBOdhyBv9F0EEp2P/OdedNjbdEBgAAwWXQBL4EAAAAvwEAAACJ+E2J8+ioAAAAqAF0Awl90NHn/8516+sERIlV0P9F0EyJ74tNzEiJ+IPBAkgrRSBIOUXQd1RIif5IK3XQSItVGKyqSDnXcwT/yXX1SYn9D7bwTDttGA+C9fz//+gwAAAAKcBIi1UQTCtlCESJIkiLVWBMK20gRIkqSIPEKEFcQV1BXUFfW13DXli4AQAAAOvSgfsAAAABcgHDweMITDtlAHPmQcHgCEWKBCRJg8QBwynATY0cQ4H7AAAAAXMVweMITDtlAHPBQcHgCEWKBCRJg8QBidlBD7cTwekLD6/KQTnIcxOJy7kACAAAKdHB6QVmQQELAcDDKcvB6gVBKchmQSkTAcCDwAHDSLj////////////gbXN2Y3J0LmRsbHxtYWxsb2MASLj////////////gZnJlZQA='))
		Else
			$CodeBase = _BinaryCall_Create(_BinaryCall_Base64Decode('VYnlVzH/VlOD7EyLXQiKC4D54A+HxQAAADHA6wWD6S2I0ID5LI1QAXfziEXmMcDrBYPpCYjQgPkIjVABd/OIReWLRRSITeSLUwkPtsmLcwWJEA+2ReUBwbgAAwAA0+CNhABwDgAAiQQk6EcEAACJNCSJRdToPAQAAItV1InHi0Xkhf+JArgBAAAAdDaF9nQyi0UQg8MNiRQkiXQkFIl8JBCJRCQYjUXgiUQkDItFDIlcJASD6A2JRCQI6CkAAACLVdSJRdSJFCToAQQAAItF1IXAdAqJPCQx/+jwAwAAg8RMifhbXl9dw1dWU1WJ5YtFJAFFKFD8i3UYAXUcVot1FK2SUopO/oPI/9Pg99BQiPGDyP/T4PfQUADRifeD7AwpwEBQUFBQUFcp9laDy/+4AAMAANPgjYg2BwAAuAAEAATR6fOragVZ6MoCAADi+Yt9/ItF8Ct9JCH4iUXosADoywIAAA+FhQAAAIpN9CN97NPngOkI9tnT7lgB916NPH/B5wg8B1qNjH5sDgAAUVa+AAEAAFCwAXI0i338K33cD7Y/i23M0eeJ8SH+AfGNbE0A6JgCAACJwcHuCIPhATnOvgABAAB1DjnwctfrDIttzOh5AgAAOfBy9FqD+gSJ0XIJg/oKsQNyArEGKcpS60mwwOhJAgAAdRRYX1pZWln/NCRRUrpkBgAAsQDrb7DM6CwCAAB1LLDw6BMCAAB1U1g8B7AJcgKwC1CLdfwrddw7dSQPgs8BAACsi338qumOAQAAsNjo9wEAAIt12HQbsOTo6wEAAIt11HQJi3XQi03UiU3Qi03YiU3Ui03ciU3YiXXcWF9ZumgKAACxCAH6Ulc8B4jIcgIEA1CLbczovAEAAHUUi0Xoi33MweADKclqCF6NfEcE6zWLbcyDxQLomwEAAHUYi0Xoi33MweADaghZaghejbxHBAEAAOsQvwQCAAADfcxqEFm+AAEAAIlN5CnAQIn96GYBAACJwSnxcvMBTeSDfcQED4OwAAAAg0XEB4tN5IP5BHIDagNZi33IweEGKcBAakBejbxPYAMAAIn96CoBAACJwSnxcvOJTeiJTdyD+QRyc4nOg2XcAdHug03cAk6D+Q5zGbivAgAAKciJ8dJl3ANF3NHgA0XIiUXM6y2D7gToowAAANHr0WXcOV3gcgb/RdwpXeBOdei4RAYAAANFyIlFzMFl3ARqBF4p/0eJ+IttzOi0AAAAqAF0Awl93NHnTnXs6wD/RdyLTeSDwQKLffyJ+CtFJDlF3HdIif4rddyLVSisqjnXcwNJdfeJffwPtvA7fSgPgnH9///oKAAAACnAjWwkPItVIIt1+Ct1GIkyi1Usi338K30kiTrJW15fw15YKcBA69qB+wAAAAFyAcPB4whWi3X4O3Ucc+SLReDB4AisiUXgiXX4XsOLTcQPtsDB4QQDRegByOsGD7bAA0XEi23IjWxFACnAjWxFAIH7AAAAAXMci0wkOMFkJCAIO0wkXHOcihH/RCQ4weMIiFQkIInZD7dVAMHpCw+vyjlMJCBzF4nLuQAIAAAp0cHpBWYBTQABwI1sJEDDweoFKUwkICnLZilVAAHAg8ABjWwkQMO4///////gbXN2Y3J0LmRsbHxtYWxsb2MAuP//////4GZyZWUA'))
		EndIf
		If Not $CodeBase Then Return SetError(1, 0, Binary(""))
	EndIf

	$Src = Binary($Src)
	Local $SrcLen = BinaryLen($Src)
	Local $SrcBuf = DllStructCreate("byte[" & $SrcLen & "]")
	DllStructSetData($SrcBuf, 1, $Src)

	Local $Ret = DllCallAddress("ptr:cdecl", $CodeBase, "ptr", DllStructGetPtr($SrcBuf), "uint_ptr", $SrcLen, "uint_ptr*", 0, "uint*", 0)
	If $Ret[0] Then
		Local $DstBuf = DllStructCreate("byte[" & $Ret[3] & "]", $Ret[0])
		Local $Output = DllStructGetData($DstBuf, 1)
		DllCall($__BinaryCall_Msvcrtdll, "none:cdecl", "free", "ptr", $Ret[0])

		Return $Output
	EndIf
	Return SetError(2, 0, Binary(""))
EndFunc

Func _BinaryCall_Relocation($Base, $Reloc)
	Local $Size = Int(BinaryMid($Reloc, 1, 2))

	For $i = 3 To BinaryLen($Reloc) Step $Size
		Local $Offset = Int(BinaryMid($Reloc, $i, $Size))
		Local $Ptr = $Base + $Offset
		DllStructSetData(DllStructCreate("ptr", $Ptr), 1, DllStructGetData(DllStructCreate("ptr", $Ptr), 1) + $Base)
	Next
EndFunc

Func _BinaryCall_ImportLibrary($Base, $Length)
	Local $JmpBin, $JmpOff, $JmpLen, $DllName, $ProcName
	If @AutoItX64 Then
		$JmpBin = Binary("0x48B8FFFFFFFFFFFFFFFFFFE0")
		$JmpOff = 2
	Else
		$JmpBin = Binary("0xB8FFFFFFFFFFE0")
		$JmpOff = 1
	EndIf
	$JmpLen = BinaryLen($JmpBin)

	Do
		Local $Ptr = _BinaryCall_MemorySearch($Base, $Length, $JmpBin)
		If $Ptr = 0 Then ExitLoop

		Local $StringPtr = $Ptr + $JmpLen
		Local $StringLen = _BinaryCall_lstrlenA($StringPtr)
		Local $String = DllStructGetData(DllStructCreate("char[" & $StringLen & "]", $StringPtr), 1)
		Local $Split = StringSplit($String, "|")

		If $Split[0] = 1 Then
			$ProcName = $Split[1]
		ElseIf $Split[0] = 2 Then
			If $Split[1] Then $DllName = $Split[1]
			$ProcName = $Split[2]
		EndIf

		If $DllName And $ProcName Then
			Local $Handle = _BinaryCall_LoadLibrary($DllName)
			If Not $Handle Then
				$__BinaryCall_LastError = "LoadLibrary fail on " & $DllName
				Return SetError(1, 0, False)
			EndIf

			Local $Proc = _BinaryCall_GetProcAddress($Handle, $ProcName)
			If Not $Proc Then
				$__BinaryCall_LastError = "GetProcAddress failed on " & $ProcName
				Return SetError(2, 0, False)
			EndIf

			DllStructSetData(DllStructCreate("ptr", $Ptr + $JmpOff), 1, $Proc)
		EndIf

		Local $Diff = Int($Ptr - $Base + $JmpLen + $StringLen + 1)
		$Base += $Diff
		$Length -= $Diff

	Until $Length <= $JmpLen
	Return True
EndFunc

Func _BinaryCall_CodePrepare($Code)
	If Not $Code Then Return ""
	If IsBinary($Code) Then Return $Code

	$Code = String($Code)
	If StringLeft($Code, 2) = "0x" Then Return Binary($Code)
	If StringIsXDigit($Code) Then Return Binary("0x" & $Code)

	Return _BinaryCall_LzmaDecompress(_BinaryCall_Base64Decode($Code))
EndFunc

Func _BinaryCall_SymbolFind($CodeBase, $Identify, $Length = Default)
	$Identify = Binary($Identify)

	If IsKeyword($Length) Then
		$Length = _BinaryCall_RegionSize($CodeBase)
	EndIf

	Local $Ptr = _BinaryCall_MemorySearch($CodeBase, $Length, $Identify)
	If $Ptr = 0 Then Return SetError(1, 0, 0)

	Return $Ptr + BinaryLen($Identify)
EndFunc

Func _BinaryCall_SymbolList($CodeBase, $Symbol)
	If Not IsArray($Symbol) Or $CodeBase = 0 Then Return SetError(1, 0, 0)

	Local $Tag = ""
	For $i = 0 To UBound($Symbol) - 1
		$Tag &=  "ptr " & $Symbol[$i] & ";"
	Next

	Local $SymbolList = DllStructCreate($Tag)
	If @Error Then Return SetError(1, 0, 0)

	For $i = 0 To UBound($Symbol) - 1
		$CodeBase = _BinaryCall_SymbolFind($CodeBase, $Symbol[$i])
		DllStructSetData($SymbolList, $Symbol[$i], $CodeBase)
	Next
	Return $SymbolList
EndFunc

Func _BinaryCall_Create($Code, $Reloc = '', $Padding = 0, $ReleaseOnExit = True, $LibraryImport = True)
	Local $BinaryCode = _BinaryCall_CodePrepare($Code)
	If Not $BinaryCode Then Return SetError(1, 0, 0)

	Local $BinaryCodeLen = BinaryLen($BinaryCode)
	Local $TotalCodeLen = $BinaryCodeLen + $Padding

	Local $CodeBase = _BinaryCall_Alloc($BinaryCode, $Padding)
	If Not $CodeBase Then Return SetError(2, 0, 0)

	If $Reloc Then
		$Reloc = _BinaryCall_CodePrepare($Reloc)
		If Not $Reloc Then Return SetError(3, 0, 0)
		_BinaryCall_Relocation($CodeBase, $Reloc)
	EndIf

	If $LibraryImport Then
		If Not _BinaryCall_ImportLibrary($CodeBase, $BinaryCodeLen) Then
			_BinaryCall_Free($CodeBase)
			Return SetError(4, 0, 0)
		EndIf
	EndIf

	If $ReleaseOnExit Then
		_BinaryCall_ReleaseOnExit($CodeBase)
	EndIf

	Return SetError(0, $TotalCodeLen, $CodeBase)
EndFunc

Func _BinaryCall_CommandLineToArgv($CommandLine, ByRef $Argc, $IsUnicode = False)
	Static $SymbolList
	If Not IsDllStruct($SymbolList) Then
		Local $Code
		If @AutoItX64 Then
			$Code = 'AwAAAASuAgAAAAAAAAAkL48ClEB9jTEOeYv4yYTosNjFNgf81Ag4vS2VP4y4wxFa+4yMI7GDB7CG+xn4JE3cdEVvk8cMp4oIuS3DgTxlcKHGVIg94tvzG/256bizZfGtAETQUCPQjW5+JSx2C/Y4C0VNJMKTlSCHiV5AzXRZ5gw3WFghbtkCCFxWOX+RDSI2oH/vROEOnqc0jfKTo17EBjqX+dW3QxrUe45xsbyYTZ9ccIGySgcOAxetbRiSxQnz8BOMbJyfrbZbuVJyGpKrXFLh/5MlBZ09Cim9qgflbGzmkrGStT9QL1f+O2krzyOzgaWWqhWL6S+y0G32RWVi0uMLR/JOGLEW/+Yg/4bzkeC0lKELT+RmWAatNa38BRfaitROMN12moRDHM6LYD1lzPLnaiefSQRVti561sxni/AFkYoCb5Lkuyw4RIn/r/flRiUg5w48YkqBBd9rXkaXrEoKwPg6rmOvOCZadu//B6HN4+Ipq5aYNuZMxSJXmxwXVRSQZVpSfLS2ATZMd9/Y7kLqrKy1H4V76SgI/d9OKApfKSbQ8ZaKIHBCsoluEip3UDOB82Z21zd933UH5l0laGWLIrTz7xVGkecjo0NQzR7LyhhoV3xszlIuw2v8q0Q/S9LxB5G6tYbOXo7lLjNIZc0derZz7DNeeeJ9dQE9hp8unubaTBpulPxTNtRjog=='
		Else
			$Code = 'AwAAAAR6AgAAAAAAAABcQfD553vjya/3DmalU0BKqABevUb/60GZ55rMwmzpQfPSRUlIl04lEiS8RDrXpS0EoBUe+uzDgZd37nVu9wsJ4fykqYvLoMz3ApxQbTBKleOIRSla6I0V8dNP3P7rHeUfjH0jCho0RvhhVpf0o4ht/iZptauxaoy1zQ19TkPZ/vf5Im8ecY6qEdHNzjo2H60jVwiOJ+1J47TmQRwxJ+yKLakq8QNxtKkRIB9B9ugfo3NAL0QslDxbyU0dSgw2aOPxV+uttLzYNnWbLBZVQbchcKgLRjC/32U3Op576sOYFolB1Nj4/33c7MRgtGLjlZfTB/4yNvd4/E+u3U6/Q4MYApCfWF4R/d9CAdiwgIjCYUkGDExKjFtHbAWXfWh9kQ7Q/GWUjsfF9BtHO6924Cy1Ou+BUKksqsxmIKP4dBjvvmz9OHc1FdtR9I63XKyYtlUnqVRtKwlNrYAZVCSFsyAefMbteq1ihU33sCsLkAnp1LRZ2wofgT1/JtT8+GO2s/n52D18wM70RH2n5uJJv8tlxQc1lwbmo4XQvcbcE91U2j9glvt2wC1pkP0hF23Nr/iiIEZHIPAOAHvhervlHE830LSHyUx8yh5Tjojr0gdLvQ=='
		EndIf
		Local $CodeBase = _BinaryCall_Create($Code)
		If @Error Then Return SetError(1, 0, 0)

		Local $Symbol[] = ["ToArgvW","ToArgvA"]
		$SymbolList = _BinaryCall_SymbolList($CodeBase, $Symbol)
		If @Error Then Return SetError(1, 0, 0)
	EndIf

	Local $Ret
	If $IsUnicode Then
		$Ret = DllCallAddress("ptr:cdecl", DllStructGetData($SymbolList, "ToArgvW"), "wstr", $CommandLine, "int*", 0)
	Else
		$Ret = DllCallAddress("ptr:cdecl", DllStructGetData($SymbolList, "ToArgvA"), "str", $CommandLine, "int*", 0)
	EndIf

	If Not @Error And $Ret[0] <> 0 Then
		_BinaryCall_ReleaseOnExit($Ret[0])
		$Argc = $Ret[2]
		Return $Ret[0]
	Else
		Return SetError(2, 0, 0)
	EndIf
EndFunc

Func _BinaryCall_StdioRedirect($Filename = "CON", $Flag = 1 + 2 + 4)
	Static $SymbolList
	If Not IsDllStruct($SymbolList) Then
		Local $Code, $Reloc
		If @AutoItX64 Then
			$Code = 'AwAAAASjAQAAAAAAAAAkL48ClEB9jTEOeYv4yYTosNjFM1rLNdMULriZUDxTj+ZdkQ01F5zKL+WDCScfQKKLn66EDmcA+gXIkPcZV4lyz8VPw8BPZlNB5KymydM15kCA+uqvmBc1V0NJfzgsF0Amhn0JhM/ZIguYCHxywMQ1SgKxUb05dxDg8WlX/2aPfSolcX47+4/72lPDNTeT7d7XRdm0ND+eCauuQcRH2YOahare9ASxuU4IMHCh2rbZYHwmTNRiQUB/8dLGtph93yhmwdHtyMPLX2x5n6sdA1mxua9htLsLTulE05LLmXbRYXylDz0A'
			$Reloc = 'AwAAAAQIAAAAAAAAAAABAB7T+CzGn9ScQAC='
		Else
			$Code = 'AwAAAASVAQAAAAAAAABcQfD553vjya/3DmalU0BKqABaUcndypZ3mTYUkHxlLV/lKZPrXYWXgNATjyiowkUQGDVYUy5THQwK4zYdU7xuGf7qfVDELc1SNbiW3NgD4D6N6ZM7auI1jPaThsPfA/ouBcx2aVQX36fjmViTZ8ZLzafjJeR7d5OG5s9sAoIzFLTZsqrFlkIJedqDAOfhA/0mMrkavTWnsio6yTbic1dER0DcEsXpLn0vBNErKHoagLzAgofHNLeFRw5yHWz5owR5CYL7rgiv2k51neHBWGx97A=='
			$Reloc = 'AwAAAAQgAAAAAAAAAAABABfyHS/VRkdjBBzbtGPD6vtmVH/IsGHYvPsTv2lGuqJxGlAA'
		EndIf

		Local $CodeBase = _BinaryCall_Create($Code, $Reloc)
		If @Error Then Return SetError(1, 0, 0)

		Local $Symbol[] = ["StdinRedirect","StdoutRedirect","StderrRedirect"]
		$SymbolList = _BinaryCall_SymbolList($CodeBase, $Symbol)
		If @Error Then Return SetError(1, 0, 0)
	EndIf

	If BitAND($Flag, 1) Then DllCallAddress("none:cdecl", DllStructGetData($SymbolList, "StdinRedirect"), "str", $Filename)
	If BitAND($Flag, 2) Then DllCallAddress("none:cdecl", DllStructGetData($SymbolList, "StdoutRedirect"), "str", $Filename)
	If BitAND($Flag, 4) Then DllCallAddress("none:cdecl", DllStructGetData($SymbolList, "StderrRedirect"), "str", $Filename)
EndFunc

Func _BinaryCall_StdinRedirect($Filename = "CON")
	Local $Ret = _BinaryCall_StdioRedirect($Filename, 1)
	Return SetError(@Error, @Extended, $Ret)
EndFunc

Func _BinaryCall_StdoutRedirect($Filename = "CON")
	Local $Ret = _BinaryCall_StdioRedirect($Filename, 2)
	Return SetError(@Error, @Extended, $Ret)
EndFunc

Func _BinaryCall_StderrRedirect($Filename = "CON")
	Local $Ret = _BinaryCall_StdioRedirect($Filename, 4)
	Return SetError(@Error, @Extended, $Ret)
EndFunc

Func _BinaryCall_ReleaseOnExit($Ptr)
	OnAutoItExitRegister('__BinaryCall_DoRelease')
	__BinaryCall_ReleaseOnExit_Handle($Ptr)
EndFunc

Func __BinaryCall_DoRelease()
	__BinaryCall_ReleaseOnExit_Handle()
EndFunc

Func __BinaryCall_ReleaseOnExit_Handle($Ptr = Default)
	Static $PtrList

	If @NumParams = 0 Then
		If IsArray($PtrList) Then
			For $i = 1 To $PtrList[0]
				_BinaryCall_Free($PtrList[$i])
			Next
		EndIf
	Else
		If Not IsArray($PtrList) Then
			Local $InitArray[1] = [0]
			$PtrList = $InitArray
		EndIf

		If IsPtr($Ptr) Then
			Local $Array = $PtrList
			Local $Size = UBound($Array)
			ReDim $Array[$Size + 1]
			$Array[$Size] = $Ptr
			$Array[0] += 1
			$PtrList = $Array
		EndIf
	EndIf
EndFunc
