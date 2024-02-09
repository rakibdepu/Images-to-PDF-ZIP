#include <Constants.au3>

#include <_CRC32ForFile.au3>

MsgBox($MB_SYSTEMMODAL, '', _IsFileDiff(@ScriptFullPath, @ScriptFullPath, 15))
MsgBox($MB_SYSTEMMODAL, '', _IsFileDiff(@ScriptFullPath, @AutoItExe, 15))

Func _IsFileDiff($sFilePath_1, $sFilePath_2, $iPercent = Default)
    Return Not (_CRC32ForFile($sFilePath_1, $iPercent) == _CRC32ForFile($sFilePath_2, $iPercent))
EndFunc   ;==>_IsFileDiff