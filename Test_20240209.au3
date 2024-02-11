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

; Open log file
Local $logFileHandle = FileOpen($logFile, $FO_APPEND) ; Append to existing log

; Write start message to log
FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Started		File comparison and synchronization." & @CRLF)

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
        FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Coping		" & $sourceFile & @CRLF)
    Else
        ; Compare file content using CRC32
        $fileHash = _IsFileDiff($sourcePath, $targetPath, 100)
        ConsoleWrite('(' & @ScriptLineNumber & ') : $fileHash = ' & $fileHash & @CRLF & '>Error code: ' & @error & @CRLF)

        If $fileHash = True Then
            ; Check filename first
            If StringInStr($sourceFile, "backup") Then
                ; Delete "backup.old"
                Local $targetDrive = "", $targetDir = "", $targetFile = "", $targetExt = ""
                Local $targetPathSplit = _PathSplit($targetPath, $targetDrive, $targetDir, $targetFile, $targetExt)
                If FileExists($targetDir & "\" & $targetFile & ".old" & $targetExt) Then
                    ConsoleWrite('(' & @ScriptLineNumber & ') : Old File = ' & $targetDir & "\" & $targetFile & ".old" & $targetExt & @CRLF & '>Error code: ' & @error & @CRLF)
                    ConsoleWrite('(' & @ScriptLineNumber & ') : Delete File = ' & $targetFile & ".old" & $targetExt & @CRLF & '>Error code: ' & @error & @CRLF)
                    FileDelete($targetFile & ".old" & $targetExt)
                EndIf
                FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Deleting		" & $targetFile & ".old" & $targetExt & @CRLF)
                FileMove($targetPath, $targetDrive & $targetDir & $targetFile & ".old" & $targetExt)
                ; Rename to "backup.old"
                FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Renaming	" & $targetFile & $targetExt & " to " & $targetFile & ".old" & $targetExt & @CRLF)
            Else
                ; Files are different, copy and log
                FileCopy($sourcePath, $targetPath)
                FileSetTime($targetPath, "", $FT_CREATED)
                FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Updating	" & $sourceFile & @CRLF)
            EndIf
        Else
            FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Same		" & $sourceFile & @CRLF)
        EndIf
    EndIf
WEnd
FileClose($fileSearch)

; Write end message to log
FileWrite($logFileHandle, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Finished	File comparison and synchronization." & @CRLF)

; Close log file
FileClose($logFileHandle)

Exit