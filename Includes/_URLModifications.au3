#include-once

#include <Array.au3>

#include "Base64.au3"
#include "_Settings.au3"

Func _ChangeImageProvider($sURL)

	If StringInStr($sURL, "bing.com/images/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(?i)(.*)(q=)", "")
		$sURL = StringRegExpReplace($sURL, "(?i)(?=&id=)(.*)", "")
		$sURL = StringReplace($sURL, " ", "+")
		Switch _GetSettingValue("Images")

			Case "Baidu"
				$sURL = "https://image.baidu.com/search/index?tn=baiduimage&word=" & $sURL

			Case "Custom"
				$sURL = _GetSettingValue("ImagePath") & $sURL

			Case "DuckDuckGo"
				$sURL = "https://duckduckgo.com/?ia=images&iax=images&q=" & $sURL

			Case "Ecosia"
				$sURL = "https://www.ecosia.org/images?q=" & $sURL

			Case "Google"
				$sURL = "https://www.google.com/search?tbm=isch&q=" & $sURL

			Case "Sogou"
				$sURL = "https://image.sogou.com/pics?query=" & $sURL

			Case "Yahoo"
				$sURL = "https://images.search.yahoo.com/search/images?p=" & $sURL

			Case "Yandex"
				$sURL = "https://yandex.com/images/search?text=" & $sURL

			Case Null
				$sURL = "https://bing.com/images/search?q=" & $sURL

			Case Else
				$sURL = _GetSettingValue("ImagePath") & $sURL

		EndSwitch
	EndIf

	Return $sURL

EndFunc

Func _ChangeNewsProvider($sURL)

	Local $sOriginal = $sURL

	Local $sRegex = "(?i).*\/(" & _
		"autos(\/enthusiasts)?" & _
		"|comics" & _
		"|companies" & _
		"|medical" & _
		"|news(\/crime|\/other|\/politics|\/us)?" & _
		"|research" & _
		"|retirement" & _
		"|sports" & _
		"|topstories" & _
		")\/"

	If StringInStr($sURL, "msn.com/") And StringRegExp($sURL, $sRegex) Then
		$sURL = StringRegExpReplace($sURL, $sRegex, "")
		$sURL = StringRegExpReplace($sURL, "(?i)(?=)\/.*", "")

		Switch _GetSettingValue("News")

			Case "DuckDuckGo"
				$sURL = "https://duckduckgo.com/?q=%5C" & $sURL & "+-site%3Amsn.com%20-site%3Abing.com"

			Case "Google"
				$sURL = "https://www.google.com/search?q=" & $sURL & "+-site%3Amsn.com%20-site%3Abing.com&btnI=I%27m+Feeling+Lucky"

			Case Null
				ContinueCase

			Case Else
				$sURL = $sOriginal

		EndSwitch
	EndIf

	Return $sURL

EndFunc

Func _ChangeSearchEngine($sURL)

	If StringInStr($sURL, "bing.com/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(?i)(.*)((\?|&)q=)", "")
		If StringInStr($sURL, "&form") Then $sURL = StringRegExpReplace($sURL, "(?i)(?=&form)(.*)", "")

		Switch _GetSettingValue("Search")

			Case "Ask"
				$sURL = "https://www.ask.com/web?q=" & $sURL

			Case "Baidu"
				$sURL = "https://www.baidu.com/s?wd=" & $sURL

			Case "Custom"
				$sURL = _GetSettingValue("SearchPath") & $sURL

			Case "DuckDuckGo"
				$sURL = "https://duckduckgo.com/?q=" & $sURL

			Case "Ecosia"
				$sURL = "https://www.ecosia.org/search?q=" & $sURL

			Case "Google"
				$sURL = "https://www.google.com/search?q=" & $sURL

			Case "Sogou"
				$sURL = "https://www.sogou.com/web?query=" & $sURL

			Case "Yahoo"
				$sURL = "https://search.yahoo.com/search?p=" & $sURL

			Case "Yandex"
				$sURL = "https://yandex.com/search/?text=" & $sURL

			Case Null
				$sURL = "https://bing.com/search?q=" & $sURL

			Case Else
				$sURL = _GetSettingValue("SearchPath") & $sURL

		EndSwitch
	EndIf

	Return $sURL

EndFunc

Func _ChangeWeatherProvider($sURL)

	;https://a.msn.com/54/en-us/ct<LATITUDE>,<LONGITUDE>?weadegreetype=F&weaext0={%22l%22:%22<CITY>%22,%22r%22:%22<STATE>%22,%22c%22:%22<COUNTRY>...
	;http://a.msn.com/54/<LOCALE>/ct<LAT>,<LONG>?ctsrc=outlook&ocid=outlookdesktopcalendar&weadegreetype=F&day=1 #126

	Local $fLat
	Local $aData
	Local $fLong
	Local $sSign
	Local $sLocale
	Local $vCoords
	Local $sOriginal = $sURL

	#forceref $sLocale

	If StringInStr($sURL, "msn.com/") Then ; TODO: Swap to Regex to reduce potential false positives

		Select ; TODO: Rewrite function. Get Provider, then Get URL Type (Forecast vs Map), Return appropriate data

			Case StringInStr($sURL, "weadegreetype")

				Select

					Case StringInStr($sURL, "/ct") ; Old Style Weather URL
						$vCoords = StringRegExpReplace($sURL, "(?i)(.*)(\/ct)", "")
						$vCoords = StringRegExpReplace($vCoords, "(?i)(?=\?weadegreetype=)(.*)", "")
						$vCoords = StringSplit($vCoords, ",")
						If $vCoords[0] = 2 Then
							$fLat = $vCoords[1]
							$fLong = $vCoords[2]
							$sSign = StringRegExpReplace($sURL, "(?i)(.*)(weadegreetype=)", "")
							$sSign = StringRegExpReplace($sSign, "(?i)(?=&weaext0=)(.*)", "")
						Else
							$sURL = $sOriginal
						EndIf

					Case StringInStr($sURL, "loc=") ; New Style Weather URL
						$vCoords = StringRegExpReplace($sURL, "(?i)(.*)(\?loc=)", "")
						$vCoords = StringRegExpReplace($vCoords, "(?i)(?=\&weadegreetype=)(.*)", "")
						$vCoords = _UnicodeURLDecode($vCoords)
						$vCoords = _Base64Decode($vCoords)
						$vCoords = BinaryToString($vCoords)
						$vCoords = StringRegExpReplace($vCoords, "(?i){|}", "")
						$aData = StringSplit($vCoords, ",")
						For $iLoop = 1 To $aData[0] Step 1
							Switch StringLeft($aData[$iLoop], 3)
								Case '"l"'
									;;;
								Case '"r"'
									;;;
								Case '"r2'
									;;;
								Case '"c"'
									;;;
								Case '"i"'
									;;;
								Case '"g"'
									$sLocale = StringTrimLeft($aData[$iLoop], 4)
									$sLocale = StringTrimRight($aData[$iLoop], 1)
								Case '"x"'
									$fLong = StringTrimLeft($aData[$iLoop], 4)
								Case '"y"'
									$fLat = StringTrimLeft($aData[$iLoop], 4)
								Case Else
									FileWrite($hLogs[$PEBIAT], _NowCalc() & " - " & "Unexpected Weather Entry: " & $aData[$iLoop] & " of " & _ArrayToString($aData) & @CRLF)
							EndSwitch
							$sSign = StringRegExpReplace($sURL, "(?i)(.*)(weadegreetype=)", "")
							$sSign = StringRegExpReplace($sSign, "(?i)(?=&weaext0=)(.*)", "")
						Next

					Case Else
						$sURL = $sOriginal

				EndSelect

				Switch _GetSettingValue("Weather")

					Case "AccuWeather"
						$sURL = "https://www.accuweather.com/en/search-locations?query=" & $fLat & "," & $fLong

					Case "Custom"
						$sURL = _GetSettingValue("WeatherPath") & $fLat & "," & $fLong

					Case "DarkSky"
						$sURL = "https://darksky.net/forecast/" & $fLat & "," & $fLong & "/"

					Case "Weather.com"
						$sURL = "https://www.weather.com/wx/today/?lat=" & $fLat & "&lon=" & $fLong & "&temp=" & $sSign ;"&locale=" & <LOCALE>

					Case "Weather.gov" ; TODO: Swap to "Government" and pass to the appropriate organization (https://en.wikipedia.org/wiki/List_of_meteorology_institutions)
						$sURL = "https://forecast.weather.gov/MapClick.php?lat=" & $fLat & "&lon=" & $fLong

					Case "Windy"
						$sURL = "https://www.windy.com/?" & $fLat & "," & $fLong

					Case "WUnderground"
						$sURL = "https://www.wunderground.com/weather/" & $fLat & "," & $fLong

					Case "Ventusky"
						$sURL = "https://www.ventusky.com/" & $fLat & ";" & $fLong

					Case "Yandex"
						$sURL = "https://yandex.ru/pogoda/?lat=" & $fLat & "&lon=" & $fLong

					Case Null
						ContinueCase

					Case Else
						$sURL = $sOriginal

				EndSwitch

			Case StringInStr($sURL, "/weather/maps")
				;;;

			Case Else
				;;;

		EndSelect

	EndIf

	Return $sURL

EndFunc

Func _ModifyURL($sURL)

	If _GetSettingValue("NoImgs") Then $sURL = _ChangeImageProvider($sURL)
	If _GetSettingValue("NoNews") Then $sURL = _ChangeNewsProvider($sURL)
	If _GetSettingValue("NoBing") Then $sURL = _ChangeSearchEngine($sURL)
	If _GetSettingValue("NoMSN") Then $sURL = _ChangeWeatherProvider($sURL)

	Return $sURL

EndFunc

Func _RedirectCMDDecode($sCMDLine)

	Local $aTemp
	Local $aCMDLine_1D
	Local $aCMDLine_2D[0][0]

	$sCMDLine = StringReplace($sCMDLine, "--edge-redirect", "Method")
	If StringInStr($sCMDLine, "https://www.bing.com/search?q=") Then ; #211
		$sCMDLine = StringReplace($sCMDLine, "&", "%26")
		$sCMDLine = StringReplace($sCMDLine, "/", "%2F")
		$sCMDLine = StringReplace($sCMDLine, "=", "%3D")
		$sCMDLine = StringReplace($sCMDLine, "Method%3D", "Method=")
	EndIf
	$sCMDLine = StringReplace($sCMDLine, "microsoft-edge:?", "&")
	$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)microsoft-edge:[\/]*", "&url=")
	$aCMDLine_1D = StringSplit($sCMDLine, "&", $STR_NOCOUNT)
	Redim $aCMDLine_2D[UBound($aCMDLine_1D)][2]
	For $iLoop = 0 To UBound($aCMDLine_1D) - 1 Step 1
		$aTemp = StringSplit($aCMDLine_1D[$iLoop], "=")
		$aCMDLine_2D[$iLoop][0] = $aTemp[1]
		If $aTemp[0] >= 2 Then $aCMDLine_2D[$iLoop][1] = $aTemp[2]
	Next

	Return $aCMDLine_2D

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _UnicodeURLDecode
; Description ...: Tranlates a URL-friendly string to a normal string
; Syntax ........: _UnicodeURLDecode($toDecode)
; Parameters ....: $toDecode           - The URL-friendly string to decode
; Return values .: The URL decoded string
; Author ........: nfwu, Dhilip89, rcmaehl
; Modified ......: 10/26/2022
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
            $i += 1
            $iOne = $aryHex[$i]
            $i += 1
            $iTwo = $aryHex[$i]
            $strChar = $strChar & Chr(Dec($iOne & $iTwo))
        Else
            $strChar = $strChar & $aryHex[$i]
        EndIf
    Next
    Local $Process = StringToBinary(StringReplace($strChar, "+", " "))
    Local $DecodedString = BinaryToString($Process, 4)
    Return $DecodedString
EndFunc   ;==>_UnicodeURLDecode