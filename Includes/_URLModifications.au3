#include-once

#include <Array.au3>

#include "Base64.au3"
#include "_Settings.au3"

Func _ChangeFeedProvider($sURL)

	If StringRegExp($sURL, "https?\:\/\/www.msn\.com\/[a-z]{2}-[a-z]{2}\/feed.*") Then

		Switch _GetSettingValue("Feed")

			Case "Ask"
				$sUrl = "https://www.ask.com/"

			Case "Baidu"
				$sURL = "https://news.baidu.com/"

			Case "Custom"
				$sURL = _GetSettingValue("FeedPath")

			Case "Google"
				$sURL = "https://news.google.com/"

			Case "Yahoo"
				$sURL = "https://news.yahoo.com/"

			Case Null
				$sURL = $sURL

			Case Else
				$sURL = _GetSettingValue("FeedPath")

		EndSwitch
	EndIf

	Return $sURL

EndFunc

Func _ChangeImageProvider($sURL)

	Local $sOriginal

	If StringInStr($sURL, "bing.com/images/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(?i)(.*)(q=)", "")
		$sURL = StringRegExpReplace($sURL, "(?i)(?=&id=)(.*)", "")
		$sURL = StringReplace($sURL, " ", "+")

		$sOriginal = $sURL

		Switch _GetSettingValue("Images")

			Case "Baidu"
				$sURL = "https://image.baidu.com/search/index?tn=baiduimage&word=" & $sURL

			Case "Brave"
				$sURL = "https://search.brave.com/?ia=images&iax=images&q=" & $sURL

			Case "Custom"
				$sURL = _GetSettingValue("ImagePath")
				If StringInStr($sURL, "%query%") Then
					$sURL = StringReplace($sURL, "%query%", $sOriginal)
				Else
					$sURL = $sURL & $sOriginal
				EndIf

			Case "DuckDuckGo"
				$sURL = "https://duckduckgo.com/?ia=images&iax=images&q=" & $sURL

			Case "Ecosia"
				$sURL = "https://www.ecosia.org/images?q=" & $sURL

			Case "Google"
				$sURL = "https://www.google.com/search?tbm=isch&q=" & $sURL

			Case "Sogou"
				$sURL = "https://image.sogou.com/pics?query=" & $sURL

			Case "StartPage"
				$sURL = "https://www.startpage.com/search?cat=images&query=" & $sURL

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

	Local $sOriginal

	Local $sRegex = "(?i).*\/(" & _
		"autos(\/enthusiasts)?" & _
		"|comics" & _
		"|companies" & _
		"|health(\/other)" & _
		"|medical" & _
		"|news(\/crime|\/other|\/politics|\/us)?" & _
		"|newsscienceandtechnology" & _
		"|research" & _
		"|retirement" & _
		"|scienceandtech" & _
		"|sports" & _
		"|techandscience" & _
		"|topstories" & _
		")\/"

	If StringInStr($sURL, "msn.com/") And StringRegExp($sURL, $sRegex) Then
		$sURL = StringRegExpReplace($sURL, $sRegex, "")
		$sURL = StringRegExpReplace($sURL, "(?i)(?=)\/.*", "")

		$sOriginal = $sURL

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

	Local $sOriginal

	If StringInStr($sURL, "bing.com/search?q=") Then
		$sURL = StringRegExpReplace($sURL, "(?i)(.*)((\?|&)q=)", "")
		If StringInStr($sURL, "&form") Then $sURL = StringRegExpReplace($sURL, "(?i)(?=&form)(.*)", "")

		$sOriginal = $sURL

		Switch _GetSettingValue("Search")

			Case "Ask"
				$sURL = "https://www.ask.com/web?q=" & $sURL

			Case "Baidu"
				$sURL = "https://www.baidu.com/s?wd=" & $sURL

			Case "Brave"
				$sURL = "https://search.brave.com/search?q=" & $sURL

			Case "Custom"
				$sURL = _GetSettingValue("SearchPath")
				If StringInStr($sURL, "%query%") Then
					$sURL = StringReplace($sURL, "%query%", $sOriginal)
				Else
					$sURL = $sURL & $sOriginal
				EndIf

			Case "DuckDuckGo"
				$sURL = "https://duckduckgo.com/?q=" & $sURL

			Case "Ecosia"
				$sURL = "https://www.ecosia.org/search?q=" & $sURL

			Case "Google"
				$sURL = "https://www.google.com/search?q=" & $sURL

			Case "Lemmy"
				$sURL = "https://search-lemmy.com/results?query=" & $sURL

			Case "Sogou"
				$sURL = "https://www.sogou.com/web?query=" & $sURL

			Case "Startpage"
				$sURL = "https://www.startpage.com/search?q=" & $sURL

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
						$vCoords = _WinAPI_UrlUnescape($vCoords)
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
									$sLocale = $aData[$iLoop]
									$sLocale = StringTrimLeft($sLocale, 5)
									$sLocale = StringTrimRight($sLocale, 1)
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
						$sURL = _GetSettingValue("WeatherPath")
						$sURL = StringReplace($sURL, "%lat%", $fLat)
						$sURL = StringReplace($sURL, "%long%", $fLong)
						$sURL = StringReplace($sURL, "%locale%", $sLocale)

					Case "Weather.com"
						$sURL = "https://weather.com/" & $sLocale & "/weather/today/l/" & $fLat & "," & $fLong

					Case "Weather.gov" ; TODO: Swap to "Government" and pass to the appropriate organization (https://en.wikipedia.org/wiki/List_of_meteorology_institutions)
						$sURL = "https://forecast.weather.gov/MapClick.php?lat=" & $fLat & "&lon=" & $fLong

					Case "Windy"
						$sLocale = StringLeft($sLocale, 2)
						$sURL = "https://www.windy.com/" & $sLocale & "/?" & $fLat & "," & $fLong

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

	If _GetSettingValue("NoFeed") Then $sURL = _ChangeFeedProvider($sURL)
	If _GetSettingValue("NoImgs") Then $sURL = _ChangeImageProvider($sURL)
	If _GetSettingValue("NoNews") Then $sURL = _ChangeNewsProvider($sURL)
	If _GetSettingValue("NoBing") Then $sURL = _ChangeSearchEngine($sURL)
	If _GetSettingValue("NoMSN") Then $sURL = _ChangeWeatherProvider($sURL)

	Return $sURL

EndFunc

Func _CMDLineDecode($sCMDLine)

	Local $aTemp
	Local $aUrlMeta
	Local $aCMDLine_1D
	Local $aCMDLine_2D[0][0]

	$sCMDLine = StringReplace($sCMDLine, "--single-argument ", "Method=Undefined&")
	$sCMDLine = StringReplace($sCMDLine, "--edge-redirect", "Method")

	If StringInStr($sCMDLine, "?url=") Or StringInStr($sCMDLine, "&url=") Then
		$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)microsoft-edge:\??[\/]*", "&")
	Else
		$sCMDLine = StringRegExpReplace($sCMDLine, "(?i)microsoft-edge:\??[\/]*", "&url=")
	EndIf
	If StringInStr($sCMDLine, "?url=") Then $sCMDLine = StringReplace($sCMDLine, "?url", "url")
	
	$sCMDLine = StringReplace($sCMDLine, "&&", "&")

	;TODO: Add url=<url> somehow if "url=" doesn't exist. Method=<whatver> screws this up. 
		;UPDATE: This broke command line flags HAHAHAHAHAHAHAHA

	If StringInStr($sCMDLine, "url=--") Then 
		$sCMDLine = StringSplit($sCMDLine, "url=", $STR_ENTIRESPLIT+$STR_NOCOUNT)[0]
		$sCMDLine = StringTrimRight($sCMDLine, 1)
	EndIf

	$aCMDLine_1D = StringSplit($sCMDLine, "&", $STR_NOCOUNT)
	Redim $aCMDLine_2D[UBound($aCMDLine_1D)][2]
	For $iLoop = 0 To UBound($aCMDLine_1D) - 1 Step 1
		$aTemp = StringSplit($aCMDLine_1D[$iLoop], "=")
		$aCMDLine_2D[$iLoop][0] = $aTemp[1]
		If $aTemp[0] >= 2 Then
			Switch $aTemp[1]
				Case "hubappsubpath"
					$aTemp[2] = _WinAPI_UrlUnescape($aTemp[2])
				Case "upn"
					$aTemp[2] = _WinAPI_UrlUnescape($aTemp[2])
				Case "url"
					If StringInStr($aTemp[2], "%2F") Then $aTemp[2] = _WinAPI_UrlUnescape($aTemp[2])
					If $aTemp[0] >= 3 Then $aTemp[2] = _ArrayToString($aTemp, "=", 2)
				Case Else
					;;;
			EndSwitch
			$aCMDLine_2D[$iLoop][1] = $aTemp[2]
		EndIf
	Next

	Return $aCMDLine_2D

EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _UnicodeURLDecode
; Description ...: Tranlates a URL-friendly string to a normal string
; Syntax ........: _UnicodeURLDecode($sData)
; Parameters ....: $sData           - The URL-friendly string to decode
; Return values .: The URL decoded string
; Author ........: nfwu, Dhilip89, rcmaehl
; Modified ......: 10/26/2022
; Remarks .......: 
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _UnicodeURLDecode($sData)
    Local $aData = StringSplit(StringReplace($sData,"+"," ",0,1),"%")
    $sData = ""
    For $i = 2 To $aData[0]
        $aData[1] &= Chr(Dec(StringLeft($aData[$i],2))) & StringTrimLeft($aData[$i],2)
    Next
    Return BinaryToString(StringToBinary($aData[1],1),4)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WinAPI_UrlUnescape
; Description ...: Tranlates a URL-friendly string to a normal string
; Syntax ........: _WinAPI_UrlUnescape($sData[, $dFlag])
; Parameters ....: $sURL            - The URL-friendly string to decode
;                  $dFlag           - [Optional] WinAPI Function parameters
; Return values .: The URL unescaped string
; Author ........: mistersquirrle, rcmaehl
; Modified ......: 2/8/2024
; Remarks .......: URL_DONT_UNESCAPE_EXTRA_INFO = 0x02000000
;                  URL_UNESCAPE_AS_UTF8         = 0x00040000 (Win 8+)                
;                  URL_UNESCAPE_INPLACE         = 0x00100000
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================

Func _WinAPI_UrlUnescape($sUrl, $dFlags = 0x00040000)

    ; https://learn.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-urlunescapew
    Local $aUrlUnescape = DllCall("Shlwapi.dll", "long", "UrlUnescapeW", _
            "wstr", $sUrl, _ ; PWSTR pszUrl - A pointer to a null-terminated string with the URL
            "wstr", "decodedUrl", _ ; PWSTR pszUnescaped - A pointer to a buffer that will receive a null-terminated string that contains the unescaped version of pszURL
            "dword*", 1024, _ ; DWORD *pcchUnescaped - The number of characters in the buffer pointed to by pszUnescaped
            "dword", $dFlags) ; DWORD dwFlags
    If @error Then
        ; ConsoleWrite('UrlUnescape error: ' & @error & ', LastErr: ' & _WinAPI_GetLastError() & ', LastMsg: ' & _WinAPI_GetLastErrorMessage() & @CRLF)
        Return SetError(@error, @extended, 0)
    EndIf

    If IsArray($aUrlUnescape) Then
		If $aUrlUnescape[2] <> "decodedUrl" Then Return $sURL
		Return $aUrlUnescape[2]
	EndIf
    
EndFunc   ;==>_WinAPI_UrlUnescape