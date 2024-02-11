#include <File.au3>
#include <Date.au3>
#include <Constants.au3>
#include "_CRC32ForFile.au3"

; Define folder paths
Global $sourceDir = "C:\Users\rakib\AppData\Roaming\Everything"
Global $targetDir = "E:\Everything Config"
Global $logFilePath = "E:\Everything Config\FileSyncLog.txt" ; Define log file path

SyncFiles()

; Function to check file difference
Func _IsFileDiff($sFilePath_1, $sFilePath_2, $iPercent = Default)
    Return Not (_CRC32ForFile($sFilePath_1, $iPercent) == _CRC32ForFile($sFilePath_2, $iPercent))
EndFunc   ;==>_IsFileDiff

Func SyncFiles()
    ; Open log file
	Local $logFile = FileOpen($logFilePath, $FO_APPEND) ; Append to existing log

    ; Write start message to log
	FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Started		File comparison and synchronization.")

    ; Loop through source files
	Local $search = FileFindFirstFile($sourceDir & "\*.*")
    If $search = -1 Then
        FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	No files	found in source directory.")
		Exit
    EndIf

    While 1
        Local $file = FileFindNextFile($search)
        If @error Then ExitLoop

        ; Skip special files and current directory
		If StringInStr($file, "backup") > 0 Or StringInStr($file, ".") = 0 Or $file = "." Or $file = ".." Then ContinueLoop

        Local $sourceFile = $sourceDir & "\" & $file
        Local $targetFile = $targetDir & "\" & $file

        ; Check if file exists in target directory
		If Not FileExists($targetFile) Then
            ; File doesn't exist, copy and log
			FileCopy($sourceFile, $targetFile)
            FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Coping		" & $file)
        Else
            ; Compare file content using CRC32
			If _IsFileDiff($sourceFile, $targetFile) Then
                Local $targetPathSplitDrive = "", $targetPathSplitDir = "", $targetPathSplitFile = "", $targetPathSplitExt = ""
                Local $targetPathSplit = _PathSplit($targetFile, $targetPathSplitDrive, $targetPathSplitDir, $targetPathSplitFile, $targetPathSplitExt)
                ; Rename ".v1" to ".v2"
                FileMove($targetDir & "\" & $targetPathSplitFile & ".v1" & $targetPathSplitExt, $targetDir & "\" & $targetPathSplitFile & ".v2" & $targetPathSplitExt, $FC_OVERWRITE)
                FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Replacing	" & "'" & $targetPathSplitFile & ".v2" & $targetPathSplitExt & "'" & " with " & "'" & $targetPathSplitFile & ".v1" & $targetPathSplitExt & "'")
                ; Rename ".v0" to ".v1"
                FileMove($targetDir & "\" & $targetPathSplitFile & $targetPathSplitExt, $targetDir & "\" & $targetPathSplitFile & ".v1" & $targetPathSplitExt, $FC_OVERWRITE)
                FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Replacing	" & "'" & $targetPathSplitFile & ".v1" & $targetPathSplitExt & "'" & " with " & "'" & $targetPathSplitFile & $targetPathSplitExt & "'")
                ; Files are different, copy and log
                FileCopy($sourceFile, $targetFile)
                FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Updating	" & $file)
            Else
                FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Same		" & $file)
            EndIf
        EndIf
    WEnd

    FileClose($Search)

    ; Write end message to log
    FileWriteLine($logFile, @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC & "	Finished	File comparison and synchronization.")

    ; Close log file
    FileClose($logFile)
EndFunc

Exit