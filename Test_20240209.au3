#include <File.au3>
#include <Date.au3>
#include <Constants.au3>
#include <_CRC32ForFile.au3>

; Define folder paths
$sourceDir = "C:\Users\rakib\AppData\Roaming\Everything\"
$targetDir = "E:\Everything Config\"
$logFile = "E:\Everything Config\FileSyncLog.txt" ; Define log file path

; Function to check file difference
Func _IsFileDiff($sFilePath_1, $sFilePath_2, $iPercent = Default)
    Return Not (_CRC32ForFile($sFilePath_1, $iPercent) == _CRC32ForFile($sFilePath_2, $iPercent))
EndFunc   ;==>_IsFileDiff

Func _GetCurrentDateTimeAs_YYYYMMDDHHMMSS()
    Local $sDateTime = ""
    Local Const $iStripAllWhitespaces = 8
    $sDateTime = _NowCalc()
    ConsoleWrite('(' & @ScriptLineNumber & ') : $sDateTime = ' & $sDateTime & @CRLF & '>Error code: ' & @error & @CRLF)

    $sDateTime = StringReplace($sDateTime, '/', '')
    $sDateTime = StringReplace($sDateTime, ':', '')

    Return StringStripWS($sDateTime, $iStripAllWhitespaces)
EndFunc

; Open log file
Local $logFileHandle = FileOpen($logFile, $FO_APPEND) ; Append to existing log

; Get current date and time
;Local $dateTime = _GetCurrentDateTimeAs_YYYYMMDDHHMMSS()

; Write start message to log
FileWrite($logFileHandle, _GetCurrentDateTimeAs_YYYYMMDDHHMMSS() & " Started file comparison and synchronization." & @CRLF)

; Loop through source files
Local $fileSearch = FileFindFirstFile($sourceDir & "*.*")
While 1
    ; Continue finding files
    Local $sourceFile = FileFindNextFile($fileSearch)
    If @error Then ExitLoop
    $sourcePath = $sourceDir & $sourceFile

    ; Skip special files and current directory
    If $sourceFile == "." Or $sourceFile == ".." Then
        ContinueLoop
    EndIf

    ; Check if file exists in target directory
    $targetPath = $targetDir & $sourceFile

    If Not FileExists($targetPath) Then
        ; File doesn't exist, copy and log
        FileCopy($sourcePath, $targetPath)
        FileSetTime($targetPath, "", $FT_CREATED)
        FileWrite($logFileHandle, _GetCurrentDateTimeAs_YYYYMMDDHHMMSS() & " Successfully copied " & $targetPath & @CRLF)
    Else
        ; Compare file content using CRC32
        $fileHash = _IsFileDiff($sourcePath, $targetPath, 100)

        If $fileHash = True Then
            ; Files are different, copy and log
            FileCopy($sourcePath, $targetPath)
            FileSetTime($targetPath, "", $FT_CREATED)
            FileWrite($logFileHandle, _GetCurrentDateTimeAs_YYYYMMDDHHMMSS() & " Successfully updated " & $targetPath & @CRLF)
        EndIf
    EndIf
WEnd
FileClose($fileSearch)

; Write end message to log
FileWrite($logFileHandle, _GetCurrentDateTimeAs_YYYYMMDDHHMMSS() & " Finished file comparison and synchronization." & @CRLF)

; Close log file
FileClose($logFileHandle)

Exit