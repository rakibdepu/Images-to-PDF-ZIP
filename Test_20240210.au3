#include <File.au3>
#include <Date.au3>
#include <Crypt.au3>
#include <Constants.au3>
#include <_CRC32ForFile.au3>

; Define folder paths
$sourceDir = "C:\Users\rakib\AppData\Roaming\Everything\"
$targetDir = "F:\Software\Everything\"

GetFileName()

Func GetFileName()
; Assign a Local variable the search handle of all files in the current directory.
Local $fileSearch = FileFindFirstFile($sourceDir & "*.*")

; Check if the search was successful, if not display a message and return False.
If $fileSearch = -1 Then
	;MsgBox($MB_SYSTEMMODAL, "", "Error: No files matched the search pattern.")
	Return False
EndIf

; Assign a Local variable the empty string which will contain the files names found.
Local $FoundFileName = ""

While 1
	$FoundFileName = FileFindNextFile($fileSearch)
	ConsoleWrite('(' & @ScriptLineNumber & ') : $FoundFileName = ' & $FoundFileName & @CRLF & '>Error code: ' & @error & @CRLF)
	; If there is no more file matching the search.
	If @error Then ExitLoop

	; Display the file name.
	Global $SelectedFile = $FoundFileName
	ConsoleWrite('(' & @ScriptLineNumber & ') : $SelectedFile = ' & $SelectedFile & @CRLF & '>Error code: ' & @error & @CRLF)
WEnd

; Close the search handle.
FileClose($fileSearch)
EndFunc   ;==>GetFileName

;Func _IsFileDiff($sFilePath_1, $sFilePath_2, $iPercent = Default)
;    Return Not (_CRC32ForFile($sFilePath_1, $iPercent) == _CRC32ForFile($sFilePath_2, $iPercent))
;EndFunc   ;==>_IsFileDiff
Exit