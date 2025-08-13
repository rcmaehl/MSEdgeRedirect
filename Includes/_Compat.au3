#include <Array.au3>

Func _DoesParentProcessWantEdge($sProcess) ; #229

    Switch $sProcess
        Case $sProcess = "BrowserPicker.exe"
            ContinueCase
        Case $sProcess = "BrowseRouter.exe"
            ContinueCase
        Case $sProcess = "BrowserSelect.exe" ; DEFUNCT. TODO: DOUBLE CHECK $aCMDLine[2] <- WHAT?
            ContinueCase
        Case $sProcess = "BrowserSelector.exe" ; DEFUNCT
            ContinueCase
        Case $sProcess = "bt.exe"
            ContinueCase
        Case $sProcess = "Hurl.exe"
            Return True
        Case Else
            Return False
    EndSwitch

EndFunc

Func _IsURLLocalHost(ByRef $aURL) ; #162

    Local $sCMDLine = _ArrayToString($aURL, " ", 2, -1)
    If StringRegExp($sCMDLine, "^([a-zA-Z][a-zA-Z0-9+\-.]*:(\/\/)?)?(127\.0\.0\.1|localhost|" & @ComputerName & ")[:\/].*") Then Return True

    Return False

EndFunc