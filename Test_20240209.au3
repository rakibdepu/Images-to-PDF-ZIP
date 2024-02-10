# AutoIt v3.3.14.3 ; 2023-07-04
#include <File.au3>
#include <Date.au3>
#include <Constants.au3>
#include <_CRC32ForFile.au3>

; Define folder paths
$sourceDir = "C:\Users\rakib\AppData\Roaming\Everything\"
$targetDir = "E:\Everything Config\"

; Function to check MD5 hash
Func _IsFileDiff($sFilePath_1, $sFilePath_2, $iPercent = Default)
    Return Not (_CRC32ForFile($sFilePath_1, $iPercent) == _CRC32ForFile($sFilePath_2, $iPercent))
EndFunc   ;==>_IsFileDiff

; Loop through source files
Local $fileSearch = FileFindFirstFile($sourceDir & "*.*")
While 1
    ; Continue finding files
    Local $sourceFile = FileFindNextFile($fileSearch)
	ConsoleWrite('(' & @ScriptLineNumber & ') : $sourceFile = ' & $sourceFile & @CRLF & '>Error code: ' & @error & @CRLF)
	If @error Then ExitLoop
    $sourcePath = $sourceDir & $sourceFile
	ConsoleWrite('(' & @ScriptLineNumber & ') : $sourcePath = ' & $sourcePath & @CRLF & '>Error code: ' & @error & @CRLF)

    ; Skip special files and current directory
    If $sourceFile == "." Or $sourceFile == ".." Then
        ContinueLoop
    EndIf

    ; Check if file exists in target directory
    $targetPath = $targetDir & $sourceFile
	ConsoleWrite('(' & @ScriptLineNumber & ') : $targetPath = ' & $targetPath & @CRLF & '>Error code: ' & @error & @CRLF)

    If Not FileExists($targetPath) Then
        ; File doesn't exist, copy and log
        FileCopy($sourcePath, $targetPath)
		FileSetTime($targetPath, "", $FT_CREATED)
    Else
        ; Compare file content using MD5 hash
		$fileHash = _IsFileDiff($sourcePath, $targetPath, 100)
		ConsoleWrite('(' & @ScriptLineNumber & ') : $fileHash = ' & $fileHash & @CRLF & '>Error code: ' & @error & @CRLF)

        If $fileHash = True Then
            ; Files are different, copy
            FileCopy($sourcePath, $targetPath)
			FileSetTime($targetPath, "", $FT_CREATED)
        EndIf
    EndIf
WEnd
FileClose($fileSearch)

Exit