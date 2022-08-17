; ============================================================================================================================
; File		: Json.au3 (2021.11.20)
; Purpose	: A Non-Strict JavaScript Object Notation (JSON) Parser UDF
; Author	: Ward
; Dependency: BinaryCall.au3
; Website	: http://www.json.org/index.html
;
; Source	: jsmn.c
; Author	: zserge
; Website	: http://zserge.com/jsmn.html
;
; Source	: json_string_encode.c, json_string_decode.c
; Author	: Ward
;             Jos - Added Json_Dump()
;             TheXMan - Json_ObjGetItems and some Json_Dump Fixes.
;             Jos - Changed Json_ObjGet() and Json_ObjExists() to allow for multilevel object in string.
; ============================================================================================================================

; ============================================================================================================================
; Public Functions:
;   Json_StringEncode($String, $Option = 0)
;   Json_StringDecode($String)
;   Json_IsObject(ByRef $Object)
;   Json_IsNull(ByRef $Null)
;   Json_Encode($Data, $Option = 0, $Indent = Default, $ArraySep = Default, $ObjectSep = Default, $ColonSep = Default)
;   Json_Decode($Json, $InitTokenCount = 1000)
;   Json_ObjCreate()
;   Json_ObjPut(ByRef $Object, $Key, $Value)
;   Json_ObjGet(ByRef $Object, $Key)
;   Json_ObjDelete(ByRef $Object, $Key)
;   Json_ObjExists(ByRef $Object, $Key)
;   Json_ObjGetCount(ByRef $Object)
;   Json_ObjGetKeys(ByRef $Object)
;   Json_ObjGetItems(ByRef $Object)
;   Json_ObjClear(ByRef $Object)
;   Json_Put(ByRef $Var, $Notation, $Data, $CheckExists = False)
;   Json_Get(ByRef $Var, $Notation)
;   Json_Dump($String)
; ============================================================================================================================

#include-once
#include "BinaryCall.au3"

; The following constants can be combined to form options for Json_Encode()
Global Const $JSON_UNESCAPED_UNICODE = 1 ; Encode multibyte Unicode characters literally
Global Const $JSON_UNESCAPED_SLASHES = 2 ; Don't escape /
Global Const $JSON_HEX_TAG = 4 ; All < and > are converted to \u003C and \u003E
Global Const $JSON_HEX_AMP = 8 ; All &s are converted to \u0026
Global Const $JSON_HEX_APOS = 16 ; All ' are converted to \u0027
Global Const $JSON_HEX_QUOT = 32 ; All " are converted to \u0022
Global Const $JSON_UNESCAPED_ASCII = 64 ; Don't escape ascii charcters between chr(1) ~ chr(0x1f)
Global Const $JSON_PRETTY_PRINT = 128 ; Use whitespace in returned data to format it
Global Const $JSON_STRICT_PRINT = 256 ; Make sure returned JSON string is RFC4627 compliant
Global Const $JSON_UNQUOTED_STRING = 512 ; Output unquoted string if possible (conflicting with $Json_STRICT_PRINT)

; Error value returnd by Json_Decode()
Global Const $JSMN_ERROR_NOMEM = -1 ; Not enough tokens were provided
Global Const $JSMN_ERROR_INVAL = -2 ; Invalid character inside JSON string
Global Const $JSMN_ERROR_PART = -3 ; The string is not a full JSON packet, more bytes expected
Global $Total_JSON_DUMP_Output = ""

Func __Jsmn_RuntimeLoader($ProcName = "")
	Static $SymbolList
	If Not IsDllStruct($SymbolList) Then
		Local $Code
		If @AutoItX64 Then
			$Code = 'AwAAAAQfCAAAAAAAAAA1HbEvgTNrvX54gCiWSTVmt5v7RCdoFJ/zhkKmwcm8yVqZPjJBoVhNHHAIzrHWKbZh1J0QAUaHB5zyQTilTmWa9O0OKeLrk/Jg+o7CmMzjEk74uPongdHv37nwYXvg97fiHvjP2bBzI9gxSkKq9Cqh/GxSHIlZPYyW76pXUt//25Aqs2Icfpyay/NFd50rW7eMliH5ynkrp16HM1afithVrO+LpSaz/IojowApmXnBHUncHliDqbkx6/AODUkyDm1hj+AiEZ9Me1Jy+hBQ1/wC/YnuuYSJvNAKp6XDnyc8Nwr54Uqx5SbUW2CezwQQ7aXX/HFiHSKpQcFW/gi8oSx5nsoxUXVjxeNI/L7z6GF2mfu3Tnpt7hliWEdA2r2VB+TIM7Pgwl9X3Ge0T3KJQUaRtLJZcPvVtOuKXr2Q9wy7hl80hVRrt9zYrbjBHXLrRx/HeIMkZwxhmKo/dD/vvaNgE+BdU8eeJqFBJK2alrK2rh2WkRynftyepm1WrdKrz/5KhQPp/4PqH+9IADDjoGBbfvJQXdT+yiO8DtfrVnd+JOEKsKEsdgeM3UXx5r6tEHO9rYWbzbnyEiX7WozZemry+vBZMMtHn1aA63+RcDQED73xOsnj00/9E5Z6hszM5Hi8vi6Hw3iOgf3cHwcXG44aau0JpuA2DlrUvnJOYkNnY+bECeSdAR1UQkFNyqRoH2xm4Y7gYMCPsFtPBlwwleEKI27SsUq1ZHVQvFCoef7DXgf/GwPCAvwDMIQfb3hJtIVubOkASRQZVNIJ/y4KPrn/gcASV7fvMjE34loltTVlyqprUWxpI51tN6vhTOLAp+CHseKxWaf9g1wdbVs0e/5xAiqgJbmKNi9OYbhV/blpp3SL63XKxGiHdxhK1aR+4rUY4eckNbaHfW7ob+q7aBoHSs6LVX9lWakb/xWxwQdwcX/7/C+TcQSOOg6rLoWZ8wur9qp+QwzoCbXkf04OYpvD5kqgEiwQnB90kLtcA+2XSbDRu+aq02eNNCzgkZujeL/HjVISjf2EuQKSsZkBhS15eiXoRgPaUoQ5586VS7t7rhM8ng5LiVzoUQIZ0pNKxWWqD+gXRBvOMIXY2yd0Ei4sE5KFIEhbs3u8vwP7nFLIpZ/RembPTuc0ZlguGJgJ2F5iApfia+C2tRYRNjVCqECCveWw6P2Btfaq9gw7cWWmJflIQbjxtccDqsn52cftLqXSna9zk05mYdJSV8z2W7vM1YJ5Rd82v0j3kau710A/kQrN41bdaxmKjL+gvSRlOLB1bpvkCtf9+h+eVA4XIkIXKFydr1OjMZ8wq2FIxPJXskAe4YMgwQmeWZXMK1KBbLB3yQR1YOYaaHk1fNea9KsXgs5YLbiP/noAusz76oEDo/DJh1aw7cUwdhboVPg1bNq88mRb5RGa13KDK9uEET7OA02KbSL+Q4HOtyasLUoVrZzVyd8iZPoGrV36vHnj+yvG4fq6F/fkug/sBRp186yVZQVmdAgFd+WiRLnUjxHUKJ6xBbpt4FTP42E/PzPw3JlDb0UQtXTDnIL0CWqbns2E7rZ5PBwrwQYwvBn/gaEeLVGDSh84DfW4zknIneGnYDXdVEHC+ITzejAnNxb1duB+w2aVTk64iXsKHETq53GMH6DuFi0oUeEFb/xp0HsRyNC8vBjOq3Kk7NZHxCQLh7UATFttG7sH+VIqGjjNwmraGJ0C92XhpQwSgfAb3KHucCHGTTti0sn6cgS3vb36BkjGKsRhXVuoQCFH96bvTYtl8paQQW9ufRfvxPqmU0sALdR0fIvZwd7Z8z0UoEec6b1Sul4e60REj/H4scb6N2ryHBR9ua5N1YxJu1uwgoLXUL2wT9ZPBjPjySUzeqXikUIKKYgNlWy+VlNIiWWTPtKpCTr508logA=='
		Else
			$Code = 'AwAAAASFBwAAAAAAAAA1HbEvgTNrvX54gCiqsa1mt5v7RCdoAFjCfVE40DZbE5UfabA9UKuHrjqOMbvjSoB2zBJTEYEQejBREnPrXL3VwpVOW+L9SSfo0rTfA8U2W+Veqo1uy0dOsPhl7vAHbBHrvJNfEUe8TT0q2eaTX2LeWpyrFEm4I3mhDJY/E9cpWf0A78e+y4c7NxewvcVvAakIHE8Xb8fgtqCTVQj3Q1eso7n1fKQj5YsQ20A86Gy9fz8dky78raeZnhYayn0b1riSUKxGVnWja2i02OvAVM3tCCvXwcbSkHTRjuIAbMu2mXF1UpKci3i/GzPmbxo9n/3aX/jpR6UvxMZuaEDEij4yzfZv7EyK9WCNBXxMmtTp3Uv6MZsK+nopXO3C0xFzZA/zQObwP3zhJ4sdatzMhFi9GAM70R4kgMzsxQDNArueXj+UFzbCCFZ89zXs22F7Ixi0FyFTk3jhH56dBaN65S+gtPztNGzEUmtk4M8IanhQSw8xCXr0x0MPDpDFDZs3aN5TtTPYmyk3psk7OrmofCQGG5cRcqEt9902qtxQDOHumfuCPMvU+oMjzLzBVEDnBbj+tY3y1jvgGbmEJguAgfB04tSeAt/2618ksnJJK+dbBkDLxjB4xrFr3uIFFadJQWUckl5vfh4MVXbsFA1hG49lqWDa7uSuPCnOhv8Yql376I4U4gfcF8LcgorkxS+64urv2nMUq6AkBEMQ8bdkI64oKLFfO7fGxh5iMNZuLoutDn2ll3nq4rPi4kOyAtfhW0UPyjvqNtXJ/h0Wik5Mi8z7BVxaURTDk81TP8y9+tzjySB/uGfHFAzjF8DUY1vqJCgn0GQ8ANtiiElX/+Wnc9HWi2bEEXItbm4yv97QrEPvJG9nPRBKWGiAQsIA5J+WryX5NrfEfRPk0QQwyl16lpHlw6l0UMuk7S21xjQgyWo0MywfzoBWW7+t4HH9sqavvP4dYAw81BxXqVHQhefUOS23en4bFUPWE98pAN6bul+kS767vDK34yTC3lA2a8wLrBEilmFhdB74fxbAl+db91PivhwF/CR4Igxr35uLdof7+jAYyACopQzmsbHpvAAwT2lapLix8H03nztAC3fBqFSPBVdIv12lsrrDw4dfhJEzq7AbL/Y7L/nIcBsQ/3UyVnZk4kZP1KzyPCBLLIQNpCVgOLJzQuyaQ6k2QCBy0eJ0ppUyfp54LjwVg0X7bwncYbAomG4ZcFwTQnC2AX3oYG5n6Bz4SLLjxrFsY+v/SVa+GqH8uePBh1TPkHVNmzjXXymEf5jROlnd+EjfQdRyitkjPrg2HiQxxDcVhCh5J2L5+6CY9eIaYgrbd8zJnzAD8KnowHwh2bi4JLgmt7ktJ1XGizox7cWf3/Dod56KAcaIrSVw9XzYybdJCf0YRA6yrwPWXbwnzc/4+UDkmegi+AoCEMoue+cC7vnYVdmlbq/YLE/DWJX383oz2Ryq8anFrZ8jYvdoh8WI+dIugYL2SwRjmBoSwn56XIaot/QpMo3pYJIa4o8aZIZrjvB7BXO5aCDeMuZdUMT6AXGAGF1AeAWxFd2XIo1coR+OplMNDuYia8YAtnSTJ9JwGYWi2dJz3xrxsTQpBONf3yn8LVf8eH+o5eXc7lzCtHlDB+YyI8V9PyMsUPOeyvpB3rr9fDfNy263Zx33zTi5jldgP2OetUqGfbwl+0+zNYnrg64bluyIN/Awt1doDCQkCKpKXxuPaem/SyCHrKjg'
		EndIf

		Local $Symbol[] = ["jsmn_parse", "jsmn_init", "json_string_decode", "json_string_encode"]
		Local $CodeBase = _BinaryCall_Create($Code)
		If @error Then Exit MsgBox(16, "Json", "Startup Failure!")

		$SymbolList = _BinaryCall_SymbolList($CodeBase, $Symbol)
		If @error Then Exit MsgBox(16, "Json", "Startup Failure!")
	EndIf
	If $ProcName Then Return DllStructGetData($SymbolList, $ProcName)
EndFunc   ;==>__Jsmn_RuntimeLoader

Func Json_StringEncode($String, $Option = 0)
	Static $Json_StringEncode = __Jsmn_RuntimeLoader("json_string_encode")
	Local $Length = StringLen($String) * 6 + 1
	Local $Buffer = DllStructCreate("wchar[" & $Length & "]")
	Local $Ret = DllCallAddress("int:cdecl", $Json_StringEncode, "wstr", $String, "ptr", DllStructGetPtr($Buffer), "uint", $Length, "int", $Option)
	Return SetError($Ret[0], 0, DllStructGetData($Buffer, 1))
EndFunc   ;==>Json_StringEncode

Func Json_StringDecode($String)
	Static $Json_StringDecode = __Jsmn_RuntimeLoader("json_string_decode")
	Local $Length = StringLen($String) + 1
	Local $Buffer = DllStructCreate("wchar[" & $Length & "]")
	Local $Ret = DllCallAddress("int:cdecl", $Json_StringDecode, "wstr", $String, "ptr", DllStructGetPtr($Buffer), "uint", $Length)
	Return SetError($Ret[0], 0, DllStructGetData($Buffer, 1))
EndFunc   ;==>Json_StringDecode

Func Json_Decode($Json, $InitTokenCount = 1000)
	Static $Jsmn_Init = __Jsmn_RuntimeLoader("jsmn_init"), $Jsmn_Parse = __Jsmn_RuntimeLoader("jsmn_parse")
	If $Json = "" Then $Json = '""'
	Local $TokenList, $Ret
	Local $Parser = DllStructCreate("uint pos;int toknext;int toksuper")
	Do
		DllCallAddress("none:cdecl", $Jsmn_Init, "ptr", DllStructGetPtr($Parser))
		$TokenList = DllStructCreate("byte[" & ($InitTokenCount * 20) & "]")
		$Ret = DllCallAddress("int:cdecl", $Jsmn_Parse, "ptr", DllStructGetPtr($Parser), "wstr", $Json, "ptr", DllStructGetPtr($TokenList), "uint", $InitTokenCount)
		$InitTokenCount *= 2
	Until $Ret[0] <> $JSMN_ERROR_NOMEM

	Local $Next = 0
	Return SetError($Ret[0], 0, _Json_Token($Json, DllStructGetPtr($TokenList), $Next))
EndFunc   ;==>Json_Decode

Func _Json_Token(ByRef $Json, $Ptr, ByRef $Next)
	If $Next = -1 Then Return Null

	Local $Token = DllStructCreate("int;int;int;int", $Ptr + ($Next * 20))
	Local $Type = DllStructGetData($Token, 1)
	Local $Start = DllStructGetData($Token, 2)
	Local $End = DllStructGetData($Token, 3)
	Local $Size = DllStructGetData($Token, 4)
	$Next += 1

	If $Type = 0 And $Start = 0 And $End = 0 And $Size = 0 Then ; Null Item
		$Next = -1
		Return Null
	EndIf

	Switch $Type
		Case 0 ; Json_PRIMITIVE
			Local $Primitive = StringMid($Json, $Start + 1, $End - $Start)
			Switch $Primitive
				Case "true"
					Return True
				Case "false"
					Return False
				Case "null"
					Return Null
				Case Else
					If StringRegExp($Primitive, "^[+\-0-9]") Then
						Return Number($Primitive)
					Else
						Return Json_StringDecode($Primitive)
					EndIf
			EndSwitch

		Case 1 ; Json_OBJECT
			Local $Object = Json_ObjCreate()
			For $i = 0 To $Size - 1 Step 2
				Local $Key = _Json_Token($Json, $Ptr, $Next)
				Local $Value = _Json_Token($Json, $Ptr, $Next)
				If Not IsString($Key) Then $Key = Json_Encode($Key)

				If $Object.Exists($Key) Then $Object.Remove($Key)
				$Object.Add($Key, $Value)
			Next
			Return $Object

		Case 2 ; Json_ARRAY
			Local $Array[$Size]
			For $i = 0 To $Size - 1
				$Array[$i] = _Json_Token($Json, $Ptr, $Next)
			Next
			Return $Array

		Case 3 ; Json_STRING
			Return Json_StringDecode(StringMid($Json, $Start + 1, $End - $Start))
	EndSwitch
EndFunc   ;==>_Json_Token

Func Json_IsObject(ByRef $Object)
	Return (IsObj($Object) And ObjName($Object) = "Dictionary")
EndFunc   ;==>Json_IsObject

Func Json_IsNull(ByRef $Null)
	Return IsKeyword($Null) Or (Not IsObj($Null) And VarGetType($Null) = "Object")
EndFunc   ;==>Json_IsNull

Func Json_Encode_Compact($Data, $Option = 0)
	Local $Json = ""

	Select
		Case IsString($Data)
			Return '"' & Json_StringEncode($Data, $Option) & '"'

		Case IsNumber($Data)
			Return $Data

		Case IsArray($Data) And UBound($Data, 0) = 1
			$Json = "["
			For $i = 0 To UBound($Data) - 1
				$Json &= Json_Encode_Compact($Data[$i], $Option) & ","
			Next
			If StringRight($Json, 1) = "," Then $Json = StringTrimRight($Json, 1)
			Return $Json & "]"

		Case Json_IsObject($Data)
			$Json = "{"
			Local $Keys = $Data.Keys()
			For $i = 0 To UBound($Keys) - 1
				$Json &= '"' & Json_StringEncode($Keys[$i], $Option) & '":' & Json_Encode_Compact($Data.Item($Keys[$i]), $Option) & ","
			Next
			If StringRight($Json, 1) = "," Then $Json = StringTrimRight($Json, 1)
			Return $Json & "}"

		Case IsBool($Data)
			Return StringLower($Data)

		Case IsPtr($Data)
			Return Number($Data)

		Case IsBinary($Data)
			Return '"' & Json_StringEncode(BinaryToString($Data, 4), $Option) & '"'

		Case Else ; Keyword, DllStruct, Object
			Return "null"
	EndSelect
EndFunc   ;==>Json_Encode_Compact

Func Json_Encode_Pretty($Data, $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF = Default, $ObjectCRLF = Default, $NextIdent = "")
	Local $ThisIdent = $NextIdent, $Json = "", $String = "", $Match = "", $Keys = ""
	Local $Length = 0

	Select
		Case IsString($Data)
			$String = Json_StringEncode($Data, $Option)
			If BitAND($Option, $JSON_UNQUOTED_STRING) And Not BitAND($Option, $JSON_STRICT_PRINT) And Not StringRegExp($String, "[\s,:]") And Not StringRegExp($String, "^[+\-0-9]") Then
				Return $String
			Else
				Return '"' & $String & '"'
			EndIf

		Case IsArray($Data) And UBound($Data, 0) = 1
			If UBound($Data) = 0 Then Return "[]"
			If IsKeyword($ArrayCRLF) Then
				$ArrayCRLF = ""
				$Match = StringRegExp($ArraySep, "[\r\n]+$", 3)
				If IsArray($Match) Then $ArrayCRLF = $Match[0]
			EndIf

			If $ArrayCRLF Then $NextIdent &= $Indent
			$Length = UBound($Data) - 1
			For $i = 0 To $Length
				If $ArrayCRLF Then $Json &= $NextIdent
				$Json &= Json_Encode_Pretty($Data[$i], $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF, $ObjectCRLF, $NextIdent)
				If $i < $Length Then $Json &= $ArraySep
			Next

			If $ArrayCRLF Then Return "[" & $ArrayCRLF & $Json & $ArrayCRLF & $ThisIdent & "]"
			Return "[" & $Json & "]"

		Case Json_IsObject($Data)
			If $Data.Count = 0 Then Return "{}"
			If IsKeyword($ObjectCRLF) Then
				$ObjectCRLF = ""
				$Match = StringRegExp($ObjectSep, "[\r\n]+$", 3)
				If IsArray($Match) Then $ObjectCRLF = $Match[0]
			EndIf

			If $ObjectCRLF Then $NextIdent &= $Indent
			$Keys = $Data.Keys()
			$Length = UBound($Keys) - 1
			For $i = 0 To $Length
				If $ObjectCRLF Then $Json &= $NextIdent
				$Json &= Json_Encode_Pretty(String($Keys[$i]), $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep) & $ColonSep _
						 & Json_Encode_Pretty($Data.Item($Keys[$i]), $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF, $ObjectCRLF, $NextIdent)
				If $i < $Length Then $Json &= $ObjectSep
			Next

			If $ObjectCRLF Then Return "{" & $ObjectCRLF & $Json & $ObjectCRLF & $ThisIdent & "}"
			Return "{" & $Json & "}"

		Case Else
			Return Json_Encode_Compact($Data, $Option)

	EndSelect
EndFunc   ;==>Json_Encode_Pretty

Func Json_Encode($Data, $Option = 0, $Indent = Default, $ArraySep = Default, $ObjectSep = Default, $ColonSep = Default)
	If BitAND($Option, $JSON_PRETTY_PRINT) Then
		Local $Strict = BitAND($Option, $JSON_STRICT_PRINT)

		If IsKeyword($Indent) Then
			$Indent = @TAB
		Else
			$Indent = Json_StringDecode($Indent)
			If StringRegExp($Indent, "[^\t ]") Then $Indent = @TAB
		EndIf

		If IsKeyword($ArraySep) Then
			$ArraySep = "," & @CRLF
		Else
			$ArraySep = Json_StringDecode($ArraySep)
			If $ArraySep = "" Or StringRegExp($ArraySep, "[^\s,]|,.*,") Or ($Strict And Not StringRegExp($ArraySep, ",")) Then $ArraySep = "," & @CRLF
		EndIf

		If IsKeyword($ObjectSep) Then
			$ObjectSep = "," & @CRLF
		Else
			$ObjectSep = Json_StringDecode($ObjectSep)
			If $ObjectSep = "" Or StringRegExp($ObjectSep, "[^\s,]|,.*,") Or ($Strict And Not StringRegExp($ObjectSep, ",")) Then $ObjectSep = "," & @CRLF
		EndIf

		If IsKeyword($ColonSep) Then
			$ColonSep = ": "
		Else
			$ColonSep = Json_StringDecode($ColonSep)
			If $ColonSep = "" Or StringRegExp($ColonSep, "[^\s,:]|[,:].*[,:]") Or ($Strict And (StringRegExp($ColonSep, ",") Or Not StringRegExp($ColonSep, ":"))) Then $ColonSep = ": "
		EndIf

		Return Json_Encode_Pretty($Data, $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep)

	ElseIf BitAND($Option, $JSON_UNQUOTED_STRING) Then
		Return Json_Encode_Pretty($Data, $Option, "", ",", ",", ":")
	Else
		Return Json_Encode_Compact($Data, $Option)
	EndIf
EndFunc   ;==>Json_Encode

Func Json_ObjCreate()
	Local $Object = ObjCreate('Scripting.Dictionary')
	$Object.CompareMode = 0
	Return $Object
EndFunc   ;==>Json_ObjCreate

Func Json_ObjPut(ByRef $Object, $Key, $Value)
	$Key = String($Key)
	If $Object.Exists($Key) Then $Object.Remove($Key)
	$Object.Add($Key, $Value)
EndFunc   ;==>Json_ObjPut

Func Json_ObjGet(ByRef $Object, $Key)
	Local $DynObject = $Object
	Local $Keys = StringSplit($Key, ".")
	For $x = 1 To $Keys[0]
		If $DynObject.Exists($Keys[$x]) Then
			If $x = $Keys[0] Then
				Return $DynObject.Item($Keys[$x])
			Else
				$DynObject = Json_ObjGet($DynObject, $Keys[$x])
			EndIf
		EndIf
	Next
	Return SetError(1, 0, '')
EndFunc   ;==>Json_ObjGet

Func Json_ObjDelete(ByRef $Object, $Key)
	$Key = String($Key)
	If $Object.Exists($Key) Then $Object.Remove($Key)
EndFunc   ;==>Json_ObjDelete

Func Json_ObjExists(ByRef $Object, $Key)
	Local $DynObject = $Object
	Local $Keys = StringSplit($Key, ".")
	For $x = 1 To $Keys[0]
		If $DynObject.Exists($Keys[$x]) Then
			If $x = $Keys[0] Then
				Return True
			Else
				$DynObject = Json_ObjGet($DynObject, $Keys[$x])
			EndIf
		Else
			Return False
		EndIf
	Next
	Return False
EndFunc   ;==>Json_ObjExists

Func Json_ObjGetCount(ByRef $Object)
	Return $Object.Count
EndFunc   ;==>Json_ObjGetCount

Func Json_ObjGetKeys(ByRef $Object)
	Return $Object.Keys()
EndFunc   ;==>Json_ObjGetKeys

Func Json_ObjGetItems(ByRef $Object)
	Return $Object.Items()
EndFunc   ;==>Json_ObjGetItems

Func Json_ObjClear(ByRef $Object)
	Return $Object.RemoveAll()
EndFunc   ;==>Json_ObjClear

; Both dot notation and square bracket notation can be supported
Func Json_Put(ByRef $Var, $Notation, $Data, $CheckExists = False)
	;Dot-notation and bracket-notation regular expressions
	Const $REGEX_DOT_WITH_STRING      = '^\.("[^"]+")', _
	      $REGEX_DOT_WITH_LITERAL     = '^\.([^.[]+)', _
	      $REGEX_BRACKET_WITH_STRING  = '^\[("[^"]+")]', _
	      $REGEX_BRACKET_WITH_LITERAL = '^\[([^]]+)]'

	Local $Ret = 0, $Item = "", $Error = 0
	Local $Match = ""

	;Set regular expression based on identified notation type.
	;Note: The order below matters. Check "string" notations
	;before their "literal" counterpart.
	Local $Regex = ""
	Select
		Case StringRegExp($Notation, $REGEX_DOT_WITH_STRING)
			$Regex = $REGEX_DOT_WITH_STRING
		Case StringRegExp($Notation, $REGEX_DOT_WITH_LITERAL)
			$Regex = $REGEX_DOT_WITH_LITERAL
		Case StringRegExp($Notation, $REGEX_BRACKET_WITH_STRING)
			$Regex = $REGEX_BRACKET_WITH_STRING
		Case StringRegExp($Notation, $REGEX_BRACKET_WITH_LITERAL)
			$Regex = $REGEX_BRACKET_WITH_LITERAL
		Case Else
			Return SetError(2, 0, "") ; invalid notation
	EndSelect

	;Parse leading notation
	$Match = StringRegExp($Notation, $Regex, 2)
	If IsArray($Match) Then
		Local $Index
		If StringLeft($Match[0], 1) = "." Then
			$Index = String(Json_Decode($Match[1])) ;only string using dot-notation
		Else
			$Index = Json_Decode($Match[1])
		EndIf
		$Notation = StringTrimLeft($Notation, StringLen($Match[0])) ;trim leading notation

		If IsString($Index) Then
			If $CheckExists And (Not Json_IsObject($Var) Or Not Json_ObjExists($Var, $Index)) Then
				Return SetError(1, 0, False) ; no specific object
			EndIf

			If Not Json_IsObject($Var) Then $Var = Json_ObjCreate()
			If $Notation Then
				$Item = Json_ObjGet($Var, $Index)
				$Ret = Json_Put($Item, $Notation, $Data, $CheckExists)
				$Error = @error
				If Not $Error Then Json_ObjPut($Var, $Index, $Item)
				Return SetError($Error, 0, $Ret)
			Else
				Json_ObjPut($Var, $Index, $Data)
				Return True
			EndIf

		ElseIf IsInt($Index) Then
			If $Index < 0 Or ($CheckExists And (Not IsArray($Var) Or UBound($Var, 0) <> 1 Or $Index >= UBound($Var))) Then
				Return SetError(1, 0, False) ; no specific object
			EndIf

			If Not IsArray($Var) Or UBound($Var, 0) <> 1 Then
				Dim $Var[$Index + 1]
			ElseIf $Index >= UBound($Var) Then
				ReDim $Var[$Index + 1]
			EndIf

			If $Notation Then
				$Ret = Json_Put($Var[$Index], $Notation, $Data, $CheckExists)
				Return SetError(@error, 0, $Ret)
			Else
				$Var[$Index] = $Data
				Return True
			EndIf

		EndIf
	EndIf
	Return SetError(2, 0, False) ; invalid notation
EndFunc   ;==>Json_Put

; Both dot notation and square bracket notation can be supported
Func Json_Get(ByRef $Var, $Notation)
	;Dot-notation and bracket-notation regular expressions
	Const $REGEX_DOT_WITH_STRING      = '^\.("[^"]+")', _
	      $REGEX_DOT_WITH_LITERAL     = '^\.([^.[]+)', _
	      $REGEX_BRACKET_WITH_STRING  = '^\[("[^"]+")]', _
	      $REGEX_BRACKET_WITH_LITERAL = '^\[([^]]+)]'

	;Set regular expression based on identified notation type.
	;Note: The order below matters. Check "string" notations
	;before their "literal" counterpart.
	Local $Regex = ""
	Select
		Case StringRegExp($Notation, $REGEX_DOT_WITH_STRING)
			$Regex = $REGEX_DOT_WITH_STRING
		Case StringRegExp($Notation, $REGEX_DOT_WITH_LITERAL)
			$Regex = $REGEX_DOT_WITH_LITERAL
		Case StringRegExp($Notation, $REGEX_BRACKET_WITH_STRING)
			$Regex = $REGEX_BRACKET_WITH_STRING
		Case StringRegExp($Notation, $REGEX_BRACKET_WITH_LITERAL)
			$Regex = $REGEX_BRACKET_WITH_LITERAL
		Case Else
			Return SetError(2, 0, "") ; invalid notation
	EndSelect

	;Parse leading notation
	Local $Match = StringRegExp($Notation, $Regex, 2)
	If IsArray($Match) Then
		Local $Index
		If StringLeft($Match[0], 1) = "." Then
			$Index = String(Json_Decode($Match[1])) ;only string using dot-notation
		Else
			$Index = Json_Decode($Match[1])
		EndIf
		$Notation = StringTrimLeft($Notation, StringLen($Match[0])) ;trim leading notation

		Local $Item
		If IsString($Index) And Json_IsObject($Var) And Json_ObjExists($Var, $Index) Then
			$Item = Json_ObjGet($Var, $Index)
		ElseIf IsInt($Index) And IsArray($Var) And UBound($Var, 0) = 1 And $Index >= 0 And $Index < UBound($Var) Then
			$Item = $Var[$Index]
		Else
			Return SetError(1, 0, "") ; no specific object
		EndIf

		If Not $Notation Then Return $Item

		Local $Ret = Json_Get($Item, $Notation)
		Return SetError(@error, 0, $Ret)
	EndIf
EndFunc   ;==>Json_Get

; List all JSON keys and their value to the Console
Func Json_Dump($Json, $InitTokenCount = 1000)
	Static $Jsmn_Init = __Jsmn_RuntimeLoader("jsmn_init"), $Jsmn_Parse = __Jsmn_RuntimeLoader("jsmn_parse")
	If $Json = "" Then $Json = '""'
	Local $TokenList, $Ret
	$Total_JSON_DUMP_Output = ""  ; reset totaldump variable at the start of the dump process (Use for testing)
	Local $Parser = DllStructCreate("uint pos;int toknext;int toksuper")
	Do
		DllCallAddress("none:cdecl", $Jsmn_Init, "ptr", DllStructGetPtr($Parser))
		$TokenList = DllStructCreate("byte[" & ($InitTokenCount * 20) & "]")
		$Ret = DllCallAddress("int:cdecl", $Jsmn_Parse, "ptr", DllStructGetPtr($Parser), "wstr", $Json, "ptr", DllStructGetPtr($TokenList), "uint", $InitTokenCount)
		$InitTokenCount *= 2
	Until $Ret[0] <> $JSMN_ERROR_NOMEM

	Local $Next = 0
	_Json_TokenDump($Json, DllStructGetPtr($TokenList), $Next)
EndFunc   ;==>Json_Dump

Func _Json_TokenDump(ByRef $Json, $Ptr, ByRef $Next, $ObjPath = "")
	If $Next = -1 Then Return Null

	Local $Token = DllStructCreate("int;int;int;int", $Ptr + ($Next * 20))
	Local $Type = DllStructGetData($Token, 1)
	Local $Start = DllStructGetData($Token, 2)
	Local $End = DllStructGetData($Token, 3)
	Local $Size = DllStructGetData($Token, 4)
	Local $Value
	$Next += 1

	If $Type = 0 And $Start = 0 And $End = 0 And $Size = 0 Then ; Null Item
		$Next = -1
		Return Null
	EndIf

	Switch $Type
		Case 0 ; Json_PRIMITIVE
			Local $Primitive = StringMid($Json, $Start + 1, $End - $Start)
			Switch $Primitive
				Case "true"
					Return "True"
				Case "false"
					Return "False"
				Case "null"
					Return "Null"
				Case Else
					If StringRegExp($Primitive, "^[+\-0-9]") Then
						Return Number($Primitive)
					Else
						Return Json_StringDecode($Primitive)
					EndIf
			EndSwitch

		Case 1 ; Json_OBJECT
			For $i = 0 To $Size - 1 Step 2
				Local $Key = _Json_TokenDump($Json, $Ptr, $Next)
				Local $cObjPath = $ObjPath & "." & $Key
				$Value = _Json_TokenDump($Json, $Ptr, $Next, $ObjPath & "." & $Key)
				If Not (IsBool($Value) And $Value = False) Then
					If Not IsString($Key) Then
						$Key = Json_Encode($Key)
					EndIf
					; show the key and its value
					ConsoleWrite("+-> " & $cObjPath & '  =' & $Value & @CRLF)
					$Total_JSON_DUMP_Output &= "+-> " & $cObjPath & '  =' & $Value & @CRLF
				EndIf
			Next
			Return False
		Case 2 ; Json_ARRAY
			Local $sObjPath = $ObjPath
			For $i = 0 To $Size - 1
				$sObjPath = $ObjPath & "[" & $i & "]"
				$Value = _Json_TokenDump($Json, $Ptr, $Next, $sObjPath)
				If Not (IsBool($Value) And $Value = False) Then ;XC - Changed line
					; show the key and its value
					ConsoleWrite("+=> " & $sObjPath & "=>" & $Value & @CRLF)
					$Total_JSON_DUMP_Output &= "+=> " & $sObjPath & "=>" & $Value & @CRLF
				EndIf
			Next
			$ObjPath = $sObjPath
			Return False

		Case 3 ; Json_STRING
			Local $LastKey = Json_StringDecode(StringMid($Json, $Start + 1, $End - $Start))
			Return $LastKey
	EndSwitch
EndFunc   ;==>_Json_TokenDump
