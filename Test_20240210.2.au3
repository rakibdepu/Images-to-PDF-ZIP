#include-once
#include <Date.au3>

ConsoleWrite(_GetCurrentDateTimeAs_YYYYMMDDHHMMSS() & @CRLF)

Func _GetCurrentDateTimeAs_YYYYMMDDHHMMSS()
    Local Const $iStripAllWhitespaces = 8
    Local $sDateTime = _NowCalc()

    $sDateTime = StringReplace($sDateTime, '/', '')
    $sDateTime = StringReplace($sDateTime, ':', '')

    Return StringStripWS($sDateTime, $iStripAllWhitespaces)
EndFunc