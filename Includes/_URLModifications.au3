#include-once

#include <Array.au3>

#include "Base64.au3"
#include "_Settings.au3"

Func _ChangeNewsProvider($sURL)

	Local $sOriginal = $sURL

	If StringInStr($sURL, "msn.com/") And StringRegExp($sURL, ".*\/(autos(\/enthusiasts)?|comics|companies|medical|news(\/crime|\/politics|\/us)?|research|retirement|sports|topstories)\/") Then
		$sURL = StringRegExpReplace($sURL, ".*\/(autos(\/enthusiasts)?|comics|companies|medical|news(\/crime|\/politics|\/us)?|research|retirement|sports|topstories)\/", "")
		$sURL = StringRegExpReplace($sURL, "(?=)\/.*", "")
		MsgBox(0, $sOriginal, $sURL)

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
		$sURL = StringRegExpReplace($sURL, "(.*)((\?|&)q=)", "")
		If StringInStr($sURL, "&form") Then $sURL = StringRegExpReplace($sURL, "(?=&form)(.*)", "")

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
						$vCoords = StringRegExpReplace($sURL, "(.*)(\/ct)", "")
						$vCoords = StringRegExpReplace($vCoords, "(?=\?weadegreetype=)(.*)", "")
						$vCoords = StringSplit($vCoords, ",")
						If $vCoords[0] = 2 Then
							$fLat = $vCoords[1]
							$fLong = $vCoords[2]
							$sSign = StringRegExpReplace($sURL, "(.*)(weadegreetype=)", "")
							$sSign = StringRegExpReplace($sSign, "(?=&weaext0=)(.*)", "")
						Else
							$sURL = $sOriginal
						EndIf

					Case StringInStr($sURL, "loc=") ; New Style Weather URL
						$vCoords = StringRegExpReplace($sURL, "(.*)(\?loc=)", "")
						$vCoords = StringRegExpReplace($vCoords, "(?=\&weadegreetype=)(.*)", "")
						$vCoords = _UnicodeURLDecode($vCoords)
						$vCoords = _Base64Decode($vCoords)
						$vCoords = BinaryToString($vCoords)
						$vCoords = StringRegExpReplace($vCoords, "{|}", "")
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
							$sSign = StringRegExpReplace($sURL, "(.*)(weadegreetype=)", "")
							$sSign = StringRegExpReplace($sSign, "(?=&weaext0=)(.*)", "")
						Next

					Case Else
						$sURL = $sOriginal

				EndSelect

				Switch _GetSettingValue("Weather")

					Case "AccuWeather"
						$sURL = "https://www.accuweather.com/en/search-locations?query=" & $fLat & "," & $fLong

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

Func _DeEmbedImage($sURL)

	If StringInStr($sURL, "bing.com/images/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(.*)(q=)", "")
		$sURL = StringRegExpReplace($sURL, "(?=&id=)(.*)", "")
		$sURL = StringReplace($sURL, " ", "+")
		$sURL = "https://www.google.com/search?tbm=isch&q=" & $sURL
	EndIf

	Return $sURL

EndFunc

Func _ModifyURL($sURL)

	If _GetSettingValue("SrcImg") Then $sURL = _DeEmbedImage($sURL)
	If _GetSettingValue("NoBing") Then $sURL = _ChangeSearchEngine($sURL)
	If _GetSettingValue("NoMSN") Then $sURL = _ChangeWeatherProvider($sURL)
	If _GetSettingValue("NoNews") Then $sURL = _ChangeNewsProvider($sURL)

	Return $sURL

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _UnicodeURLDecode
; Description ...: Tranlates a URL-friendly string to a normal string
; Syntax ........: _UnicodeURLDecode($toDecode)
; Parameters ....: $$toDecode           - The URL-friendly string to decode
; Return values .: The URL decoded string
; Author ........: nfwu, Dhilip89, rcmaehl
; Modified ......: 12/19/2021
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