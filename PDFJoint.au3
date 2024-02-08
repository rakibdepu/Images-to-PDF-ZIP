#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icone\Full ico\documentsorcopy_V2.ico
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;(C)NSC

; PDFjoint
; nscPDFjoin
; an utility to join multiple pdf using ghostscript, sanitizing the names.
; (c) 2016-23 NSC
; V1.11 march 2020 added fileinstall of the ghostscript
; V1.2 updated chooseFileFolder and updated ghostscript.
; V2.0 complete rewrite with new system to select files
;      based on drag and drop script by user Trong in AutoIt Forum:
;      https://www.autoitscript.com/forum/topic/209558-gui-example-dragging-and-dropping-folderfiles-into-the-gui/
; V.2.1 added PDF pages extraction, output name "sanification" removing spacese and dots
; V.2.15 draggable order of files

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPIEx.au3>
#include <WinAPIMisc.au3>
#include <WinAPIInternals.au3>
#include <WinAPISysWin.au3>
#include <WinAPIShPath.au3>
#include <TrayConstants.au3>

#include <Array.au3>
#include <File.au3>

#include <GuiListView.au3>
#include <GUIListViewEx.au3>


If Not FileExists("C:\autoit\PDFjoint\gswin64c.exe") Then
    MsgBox(64, "PDFjoint - Components install", "inserting gswin64c.exe in c:\autoit\PDFjoint", 3)
    DirCreate("c:\autoit\PDFjoint")
    FileInstall("c:\NSC_test\resources\PDFJoint\gswin64c.exe", "c:\autoit\PDFjoint\gswin64c.exe", 1)
    FileInstall("c:\NSC_test\resources\PDFJoint\gsdll64.dll", "c:\autoit\PDFjoint\gsdll64.dll", 1)
    FileInstall("c:\NSC_test\resources\PDFJoint\gsdll64.lib", "c:\autoit\PDFjoint\gsdll64.lib", 1)
EndIf

Global $destfolder = "C:\autoit\PDFjoint\PDF_United"
If Not FileExists($destfolder) Then DirCreate($destfolder)

Global $ver = "V.2.15", $LWid_PDF, $LW_PDF, $aLW_PDF, $aPDF

Opt("MustDeclareVars", 1) ;0=no, 1=require pre-declaration
Opt("TrayAutoPause", 0) ;0=no pause, 1=Pause
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Global $trayShow = TrayCreateItem("Show UI")
TrayCreateItem("") ; Create a separator line.
Global $trayExit = TrayCreateItem("Exit")
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

Global Const $sGUI_Show_Title = 1
Global Const $sAppName = "-> PDF Joint ===#'" & "      " & $ver & "                NSC"
Global Const $sLabel_Title = "Drag and drop files and folders HERE !"
Global Const $sLabel_Task = "Or click the button to browse."
Global Const $sLabel_Status = "READY !"

#Region ### START GUI section ###
Global $hGUI
If $sGUI_Show_Title Then
    $hGUI = GUICreate($sAppName, 600, 500, 20, 20, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
Else
    $hGUI = GUICreate($sAppName, 600, 500, 20, 20, BitOR($WS_POPUP, $WS_BORDER), BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
EndIf
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idLabel_BG = GUICtrlCreateLabel("", 66, 0, 400, 81, -1, $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idIcon = GUICtrlCreateIcon(@WindowsDir & "\explorer.exe", -19, 1, 8, 64, 64, BitOR($GUI_SS_DEFAULT_ICON, $SS_CENTERIMAGE))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Set/UnSet Windows on TOP")
Global $idLabel_Titles = GUICtrlCreateLabel($sLabel_Title, 69, 10, 396, 17, $SS_CENTERIMAGE, $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idLabel_Task = GUICtrlCreateLabel($sLabel_Task, 69, 34, 396, 17, $SS_CENTERIMAGE, $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idLabel_Status = GUICtrlCreateLabel($sLabel_Status, 69, 58, 396, 17, BitOR($SS_CENTER, $SS_CENTERIMAGE), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_BrowseFiles = GUICtrlCreateButton("Browse &files", 472, 8, 99, 33, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

Global $idButton_About = GUICtrlCreateButton("(c)", 579, 34, 17, 17, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Show About")

Global $idProgress_Total = GUICtrlCreateProgress(1, 1, 596, 4)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idProgress_Current = GUICtrlCreateProgress(1, 75, 596, 4)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

LW_PDF_create()

GUICtrlCreateGroup("Mode:", 10, 385, 300, 40)
Global $radioJoin = GUICtrlCreateRadio("Join PDFs", 15, 400, 100, 20)
GUICtrlSetState($radioJoin, $GUI_CHECKED)
GUICtrlSetTip(-1, "mode: two or more PDF joined together")
Global $radioextr = GUICtrlCreateRadio("Extract PDF", 160, 400, 100, 20)
GUICtrlSetTip(-1, "mode: extract a new pdf from page x to page y, same extraction also for multiple PDFs")

Global $idButton_Join = GUICtrlCreateButton("Work !", 10, 430, 50, 50, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 10, 800, 80, "consolas")
GUICtrlSetTip(-1, "Join PDFs !")

Global $idButton_Join_on_Desktop = GUICtrlCreateButton("Work to Desktop", 65, 435, 100, 40, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 7, 200, 50, "consolas")
GUICtrlSetTip(-1, "Join PDFs ON DESKTOP")

Global $idButton_openfolder = GUICtrlCreateButton("Output Folder", 250, 430, 100, 50, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 10, 800, 80, "consolas")
GUICtrlSetTip(-1, "Open folder with all United PDFs")

Global $idButton_Clear = GUICtrlCreateButton("Clear List", 490, 430, 100, 50, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 10, 800, 80, "consolas")
GUICtrlSetTip(-1, "Clean the current list of PDF files")

#EndRegion ### START GUI section ###
HotKeySet('^5', '_GUI_SetOnTop')
Global $GuiOnTop = 0, $onWorking = 0
WinSetTrans($hGUI, "", 91)
_GUI_OnProgress()
_GUI_SHOW()
_GUI_OnStandby()
WinSetTrans($hGUI, "", 230)

; Allow drag and drop when run as!
_WinAPI_ChangeWindowMessageFilterEx($hGUI, $WM_DROPFILES, $MSGFLT_ALLOW) ; $WM_DROPFILES = 0x0233
_WinAPI_ChangeWindowMessageFilterEx($hGUI, $WM_COPYDATA, $MSGFLT_ALLOW) ; $WM_COPYDATA = 0x004A - $MSGFLT_ALLOW = 1
_WinAPI_ChangeWindowMessageFilterEx($hGUI, $WM_COPYGLOBALDATA, $MSGFLT_ALLOW) ; $WM_COPYGLOBALDATA = 0x0049

AdlibRegister("_GUI_ResetStatus", 5000)
AdlibRegister("_GUI_SwitchMsg", 50)
AdlibRegister("_TRAY_SwitchMsg", 50)

Global $__aDropFiles, $sPercent, $guiMsg, $trayMsg
GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES")
Global $aCmdLineRaw = StringReplace($CmdLineRaw, '/ErrorStdOut "' & @ScriptFullPath & '"', "")
;ConsoleWrite($aCmdLineRaw & @CRLF)
Global $aCmdLine = _WinAPI_CommandLineToArgv($aCmdLineRaw)
If IsArray($aCmdLine) And $aCmdLine[0] > 0 Then
    _GUI_OnProgress()
    For $i = 1 To $aCmdLine[0]
        _Main_Processing($aCmdLine[$i], $i, $aCmdLine[0])
    Next
    _GUI_OnStandby()
    GUICtrlSetData($idLabel_Status, "Everything is done!")
    Sleep(3000) ; Pause 4s
    Exit
Else
    While 1
        ;_GUI_SwitchMsg()
        Switch $guiMsg

            Case $idButton_BrowseFiles
                _GUI_OnProgress()
                ; Display an open dialog to select files.
                Local $zListFileIN, $zFileIN = FileOpenDialog("Select Files", @WorkingDir, "All File (*)", $FD_FILEMUSTEXIST + $FD_MULTISELECT, "", $hGUI) ;1+4
                If Not @error Then
                    If StringInStr($zFileIN, "|") Then
                        $zListFileIN = StringSplit($zFileIN, "|")
                        If IsArray($zListFileIN) Then
                            For $i = 2 To $zListFileIN[0]
                                _Main_Processing($zListFileIN[$i], $i - 1, $zListFileIN[0] - 1)
                            Next
                        EndIf
                    Else
                        _Main_Processing($zFileIN, 1, 1)
                    EndIf
                EndIf
                _GUI_OnStandby()

            Case $idButton_Join
                _GUI_OnProgress()
                Selector()
                LW_PDF_reset()
                $aPDF = ""

                If Not WinExists("PDF_United") Then Run("explorer.exe " & $destfolder)

                _GUI_OnStandby()

            Case $idButton_Join_on_Desktop
                _GUI_OnProgress()
                $destfolder = @DesktopDir
                Selector()
                LW_PDF_reset()
                $aPDF = ""
                $destfolder = "C:\autoit\PDFjoint\PDF_United"
                _GUI_OnStandby()

            Case $idButton_Clear
                _GUI_OnProgress()
                LW_PDF_reset()
                $aPDF = ""
                _GUI_OnStandby()

            Case $idButton_openfolder
                _GUI_OnProgress()
                Run("explorer.exe " & $destfolder)
                _GUI_OnStandby()
            Case $GUI_EVENT_DROPPED
                _GUI_OnProgress()
                If $__aDropFiles[0] > 0 Then
                    For $i = 1 To $__aDropFiles[0]
                        _Main_Processing($__aDropFiles[$i], $i, $__aDropFiles[0])
                    Next
                EndIf
                _GUI_OnStandby()
        EndSwitch
        ;_TRAY_SwitchMsg()
    WEnd
EndIf
; * -----:|
Func _Main_Processing($sFilePath, $nCurrent = 0, $nTotal = 0)
    ;_GUI_SwitchMsg()
    ;_TRAY_SwitchMsg()
    $sPercent = Round(($nCurrent / $nTotal) * 100, 2)
    GUICtrlSetData($idProgress_Total, $sPercent)
    ;ConsoleWrite("- Percent: " & $sPercent & " %" & @CRLF)
    GUICtrlSetData($idProgress_Current, 0)
    GUICtrlSetData($idLabel_Titles, "Processing " & $nCurrent & "/" & $nTotal & " folder/files ! ")

    Local $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt
    Local $aPathSplit = _SplitPath($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt)
    ;Local $sCurrentDirPath= $sDrive&$sCurrentDir;StringRegExpReplace($aPathSplit, '\\[^\\]*$', '')
    ;Local $sCurrentDirName =StringRegExpReplace(_PathRemoveBackslash($sCurrentDirPath), '.*\\', '')
    #cs
    ConsoleWrite(";~ - [1] Drive: " & $sDrive & @CRLF)
    ConsoleWrite(";~ - [2] ParentDir: " & $sParentDir & @CRLF)
    ConsoleWrite(";~ - [3] CurrentDir: " & $sCurrentDir & @CRLF)
    ConsoleWrite(";~ - [4] FileName NoExt: " & $sFileNameNoExt & @CRLF)
    ConsoleWrite(";~ - [5] Extension: " & $sExtension & @CRLF)
    ConsoleWrite(";~ - [6] FileName: " & $sFileName & @CRLF)
    ConsoleWrite(";~ - [7] PathParentDir: " & $sPathParentDir & @CRLF)
    ConsoleWrite(";~ - [8] PathCurrentDir: " & $sPathCurrentDir & @CRLF)
    ConsoleWrite(";~ - [9] PathFileName NoExt: " & $sPathFileNameNoExt & @CRLF)
    ConsoleWrite("- Processing (" & $nCurrent & "/" & $nTotal & "): " & $sFilePath & @CRLF)
    #ce
    If _IsFile($sFilePath) Then
        ;   ConsoleWrite("- Processing file: " & $sFileName & @CRLF)
        GUICtrlSetData($idLabel_Task, "Currently File: " & $sFileName)
        ; Your file handler is here!
        If Not IsArray($aPDF) Then Global $aPDF[1][3]
        $aPDF[0][0] = $aPDF[0][0] + 1
        _ArrayAdd($aPDF, $sDrive & $sParentDir & $sCurrentDir & "|" & $sFileNameNoExt & "|" & $sExtension)
        LW_PDF_reset()
        LW_PDF_populate()
    Else
        If ($sParentDir == "\" And $sCurrentDir == "") Then ; Is Root Drive
            ; Your drive handler is here!
            ;   ConsoleWrite("- Processing drive: " & $sDrive & @CRLF)
            GUICtrlSetData($idLabel_Task, "Currently Drive: " & $sDrive)
        Else
            ; Your directory handler is here!
            ;   ConsoleWrite("- Processing directory: " & _PathRemove_Backslash($sPathCurrentDir) & @CRLF)
            GUICtrlSetData($idLabel_Task, "Currently Folder: " & _PathRemove_Backslash($sCurrentDir))
        EndIf
    EndIf

    ; Code section for GUI testing only
    GUICtrlSetData($idProgress_Current, 40)
    Sleep(100) ; test gui
    GUICtrlSetData($idProgress_Current, 60)
    Sleep(100) ; test gui
    GUICtrlSetData($idProgress_Current, 80)
    Sleep(100) ; test gui
    GUICtrlSetData($idProgress_Current, 100)
    GUICtrlSetData($idLabel_Status, "Everything is done!")
    Sleep(600) ; test gui
    ; End code test GUI

EndFunc   ;==>_Main_Processing

Func _Exit()
    If $onWorking Then
        Local $IdOfButtonPressed = MsgBox($MB_ICONQUESTION + $MB_OKCANCEL + $MB_TOPMOST, "Program is working!", "Are you sure you want to exit the program?" & @CRLF & "Select [OK] to Exit - Select [Cancel] continue script", 10, $hGUI)
        If ($IdOfButtonPressed = $IDOK) Then Exit
    Else
        Exit
    EndIf
EndFunc   ;==>_Exit
Func _GUI_SwitchMsg()
    $guiMsg = GUIGetMsg()
    Switch $guiMsg
        Case $idButton_About
            MsgBox(64, $sAppName & " : © NSC 2023 ", $ver & " based on Ghostscript (AGPL version) and on drag and drop script by Ðào Van Trong - Trong.LIVE", Default, $hGUI)
            ; Case $idButton_Minimizes
            ;   GUISetState(@SW_MINIMIZE, $hGUI)
        Case $idIcon
            _GUI_SetOnTop()
        Case $GUI_EVENT_CLOSE ;, $idButton_Close
            _Exit()
    EndSwitch
EndFunc   ;==>_GUI_SwitchMsg
Func _GUI_SetOnTop()
    _GUI_SHOW()
    If $GuiOnTop Then
        $GuiOnTop = 0
        GUICtrlSetData($idLabel_Status, "The window is now normal, no longer always showing on top.")
    Else
        $GuiOnTop = 1
        GUICtrlSetData($idLabel_Status, "Set window to always show on top.")
    EndIf
    WinSetOnTop($hGUI, "", $GuiOnTop)
EndFunc   ;==>_GUI_SetOnTop
Func _TRAY_SwitchMsg()
    $trayMsg = TrayGetMsg()
    Switch $trayMsg
        Case $trayShow
            _GUI_SHOW()
        Case $trayExit
            _Exit()
    EndSwitch
EndFunc   ;==>_TRAY_SwitchMsg
Func _GUI_ResetStatus()
    ;_GUI_SHOW()
    If $onWorking Then
        GUICtrlSetData($idLabel_Status, "Working...")
    Else
        GUICtrlSetData($idLabel_Status, $sLabel_Status)
    EndIf
EndFunc   ;==>_GUI_ResetStatus
Func _GUI_SHOW()
    _WinAPI_ShowWindow(@SW_SHOW, $hGUI)
    GUISetState(@SW_SHOW, $hGUI)
    GUISetState(@SW_UNLOCK, $hGUI)
    GUISetState(@SW_ENABLE, $hGUI)
    GUISetState(@SW_RESTORE, $hGUI)
    GUISetState(@SW_SHOWNORMAL, $hGUI)
    WinActivate($hGUI)
EndFunc   ;==>_GUI_SHOW
Func _GUI_OnProgress()
    $onWorking = 1
    GUICtrlSetData($idLabel_Status, "Working...")
    GUICtrlSetData($idProgress_Total, 0)
    GUICtrlSetData($idProgress_Current, 0)
    GUICtrlSetState($idButton_BrowseFiles, $GUI_DISABLE)     ; DISABLE

    GUICtrlSetState($idButton_Join, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_Clear, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_openfolder, $GUI_NODROPACCEPTED) ; NODROPACCEPTED

    GUICtrlSetState($idLabel_BG, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idIcon, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Titles, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Task, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Status, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_BrowseFiles, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_About, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idProgress_Total, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idProgress_Current, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
EndFunc   ;==>_GUI_OnProgress
Func _GUI_OnStandby()
    $onWorking = 0
    GUICtrlSetData($idLabel_Titles, $sLabel_Title)
    GUICtrlSetData($idLabel_Task, $sLabel_Task)
    GUICtrlSetData($idLabel_Status, $sLabel_Status)
    GUICtrlSetData($idProgress_Total, 100)
    GUICtrlSetData($idProgress_Current, 100)
    GUICtrlSetState($idButton_BrowseFiles, $GUI_ENABLE)     ; ENABLE

    GUICtrlSetState($idButton_Join, $GUI_DROPACCEPTED)     ; DROPACCEPTED
    GUICtrlSetState($idButton_Clear, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_openfolder, $GUI_DROPACCEPTED) ; DROPACCEPTED

    GUICtrlSetState($idLabel_BG, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idIcon, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Titles, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Task, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Status, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_BrowseFiles, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_About, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idProgress_Total, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idProgress_Current, $GUI_DROPACCEPTED) ; DROPACCEPTED
EndFunc   ;==>_GUI_OnStandby
Func WM_DROPFILES($hWnd, $iMsg, $iwParam, $ilParam)
    #forceref $hWnd, $ilParam
    Switch $iMsg
        Case $WM_DROPFILES
            Local $aReturn = _WinAPI_DragQueryFileEx($iwParam)
            If IsArray($aReturn) Then
                $__aDropFiles = $aReturn
            Else
                Local $aError[1] = [0]
                $__aDropFiles = $aError
            EndIf
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DROPFILES
; * -----:|
Func _IsFile($sPath)
    If (Not FileExists($sPath)) Then Return SetError(-1, 0, 0)
    If StringInStr(FileGetAttrib($sPath), 'D') <> 0 Then
        Return SetError(0, 0, 0)
    Else
        Return SetError(0, 0, 1)
    EndIf
EndFunc   ;==>_IsFile

#Region listview

Func LW_PDF_create() ;crea listview Eventi

    $LW_PDF = GUICtrlCreateListView("PDF files path|PDF files|ext ", 2, 82, 596, 300, BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_SHOWSELALWAYS, $WS_BORDER))
    _GUICtrlListView_SetExtendedListViewStyle($LW_PDF, $LVS_EX_FULLROWSELECT)

    _GUICtrlListView_SetColumnWidth($LW_PDF, 0, 300)
    _GUICtrlListView_SetColumnWidth($LW_PDF, 1, 250)
    _GUICtrlListView_SetColumnWidth($LW_PDF, 2, 40)
    GUICtrlSetFont($LW_PDF, 9, 800, 0, "verdana")

    GUICtrlSetResizing($LW_PDF, 102)

    $LWid_PDF = _GUIListViewEx_Init($LW_PDF, "", 0, Default, False, 33)
    _GUIListViewEx_MsgRegister() ; If you do not do this they UDF will not work at all <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    GUISetState()

EndFunc   ;==>LW_PDF_create

Func LW_PDF_reset() ; cancella e ricrea

    If IsArray($aLW_PDF) Then
        If $aLW_PDF[0][0] <> 0 Then

            _GUIListViewEx_DeleteSpec($LWid_PDF, LW_RangeGenerator($aLW_PDF[0][0]))
        EndIf
    EndIf
EndFunc   ;==>LW_PDF_reset

Func LW_PDF_populate()

    _GUIListViewEx_BlockReDraw($LWid_PDF, True) ; perla pearl anti flickering ( use with same instruction with FALSE to close, look down...)

    Local $visibileRows = 0

    Local $imeno1
    For $i = 1 To UBound($aPDF) - 1
        $imeno1 = $i - 1

        Local $list_fields = _ArrayToString($aPDF, "|", $i, $i)     ; possible pearl

        _GUIListViewEx_InsertSpec($LWid_PDF, $imeno1, $list_fields, False, False)
        If @error Then
            MsgBox($MB_OK, @error, "_GUIListViewEx_InsertSpec(" & $i & ") Failed -- Aborting")
            Exit
        EndIf
        $visibileRows += 1
    Next

    _GUICtrlListView_EnsureVisible($LW_PDF, 0) ;

    _GUIListViewEx_BlockReDraw($LWid_PDF, False)

    $aLW_PDF = _GUIListViewEx_ReadToArray($LW_PDF, 1)


EndFunc   ;==>LW_PDF_populate


Func LW_RangeGenerator($range)
    Local $text

    For $i = 0 To $range - 1

        If $i = 0 Then
            $text = $i
        Else
            $text = $text & ";" & $i
        EndIf

    Next

    Return $text

EndFunc   ;==>LW_RangeGenerator

#EndRegion listview

Func _SplitPath($sFilePath, ByRef $sDrive, ByRef $sParentDir, ByRef $sCurrentDir, ByRef $sFileNameNoExt, ByRef $sExtension, ByRef $sFileName, ByRef $sPathParentDir, ByRef $sPathCurrentDir, ByRef $sPathFileNameNoExt)

    $sFilePath = _PathFix($sFilePath)
    _PathSplit_Ex($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension)
    $sFileName = $sFileNameNoExt & $sExtension
    $sPathParentDir = $sDrive & $sParentDir
    $sPathCurrentDir = $sDrive & $sParentDir & $sCurrentDir
    $sPathFileNameNoExt = $sDrive & $sParentDir & $sCurrentDir & $sFileNameNoExt
    Local $aSplitPath[10] = [$sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt]
    Return $aSplitPath
EndFunc   ;==>_SplitPath
; * -----:|
Func _PathSplit_Ex($sFilePath, ByRef $sDrive, ByRef $sParentDir, ByRef $sCurrentDir, ByRef $sFileNameNoExt, ByRef $sExtension)
    $sFilePath = _PathFix($sFilePath)
    $sFilePath = StringRegExp($sFilePath & " ", "^((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*?[\/\\]+)?([^\/\\]*[\/\\])?[\/\\]*((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", 1)
    $sDrive = (StringStripWS($sFilePath[1], 8) == "") ? "" : $sFilePath[1]
    $sFilePath[2] = StringRegExpReplace($sFilePath[2], "[\/\\]+\h*", "\" & StringLeft($sFilePath[2], 1))
    $sParentDir = (StringStripWS($sFilePath[2], 8) == "") ? "" : $sFilePath[2]
    $sCurrentDir = (StringStripWS($sFilePath[3], 8) == "") ? "" : $sFilePath[3]
    $sFileNameNoExt = (StringStripWS($sFilePath[4], 8) == "") ? "" : $sFilePath[4]
    $sExtension = (StringStripWS($sFilePath[5], 8) == "") ? "" : StringStripWS($sFilePath[5], 3)
    Return $sFilePath
EndFunc   ;==>_PathSplit_Ex
; * -----:|
Func _PathFix($sFilePath)
    $sFilePath = StringStripWS($sFilePath, 3)
    $sFilePath = StringReplace($sFilePath, "/", "\")
    While StringInStr($sFilePath, " \")
        $sFilePath = StringReplace($sFilePath, " /", "\")
    WEnd
    While StringInStr($sFilePath, "\ ")
        $sFilePath = StringReplace($sFilePath, "/ ", "\")
    WEnd
    If (FileExists($sFilePath) And StringInStr(FileGetAttrib($sFilePath), 'D')) Then
        $sFilePath = _PathRemove_Backslash($sFilePath) & "\"
    EndIf
    Return $sFilePath
EndFunc   ;==>_PathFix
; * -----:|
Func _PathRemove_Backslash($sPath)
    If StringRight($sPath, 1) == '\' Then
        $sPath = StringTrimRight($sPath, 1)
    EndIf
    Return $sPath
EndFunc   ;==>_PathRemove_Backslash
; * -----:|


Func __WinAPI_GetLastError(Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
    Local $aCall = DllCall("kernel32.dll", "dword", "GetLastError")
    Return SetError($_iCallerError, $_iCallerExtended, $aCall[0])
EndFunc   ;==>__WinAPI_GetLastError
Func __WinAPI_CreateFileEx($sFilePath, $iCreation, $iAccess = 0, $iShare = 0, $iFlagsAndAttributes = 0, $tSecurity = 0, $hTemplate = 0)
    Local $aCall = DllCall('kernel32.dll', 'handle', 'CreateFileW', 'wstr', $sFilePath, 'dword', $iAccess, 'dword', $iShare, 'struct*', $tSecurity, 'dword', $iCreation, 'dword', $iFlagsAndAttributes, 'handle', $hTemplate)
    If @error Then Return SetError(@error, @extended, 0)
    If $aCall[0] = Ptr(-1) Then Return SetError(10, __WinAPI_GetLastError(), 0)
    Return $aCall[0]
EndFunc   ;==>__WinAPI_CreateFileEx
Func __WinAPI_CloseHandle($hObject)
    Local $aCall = DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hObject)
    If @error Then Return SetError(@error, @extended, False)
    Return $aCall[0]
EndFunc   ;==>__WinAPI_CloseHandle
; * -----:|



#Region PDF JOIner

Func Selector()

    $aLW_PDF = _GUIListViewEx_ReadToArray($LW_PDF, 1)
    $aPDF = $aLW_PDF ;V2.15

    If GUICtrlRead($radioJoin) = $GUI_CHECKED Then
        PREJoiner()

    Else
        PREExtractor()

    EndIf

EndFunc   ;==>Selector

Func PREJoiner()

    Local $uno, $due, $tre, $nomefinale
    If IsArray($aPDF) Then

        For $i = 1 To $aPDF[0][0]

            If $i = 1 Then ; primo giro chiedo il nome del pdf finale
                $uno = filenamesanitizer($aPDF[$i][0] & $aPDF[$i][1] & $aPDF[$i][2])

                $nomefinale = InputBox("PDFjoint -  Joined PDF name ", "Write down filename without '.pdf' ", "UnitedPDF" & @MSEC)

                $nomefinale = StringStripCR($nomefinale)
                $nomefinale = StringStripWS($nomefinale, 8)
                $nomefinale = StringReplace($nomefinale, ".", "")
                $nomefinale = $nomefinale & ".pdf"

            EndIf

            If $i = 2 Then
                $due = filenamesanitizer($aPDF[$i][0] & $aPDF[$i][1] & $aPDF[$i][2])
                $tre = $destfolder & "\PDFjoint_work" & $i & ".pdf"
                joiner($uno, $due, $tre)
            EndIf

            If $i > 2 Then
                $uno = $tre
                $due = filenamesanitizer($aPDF[$i][0] & $aPDF[$i][1] & $aPDF[$i][2])
                $tre = $destfolder & "\PDFjoint_work" & $i & ".pdf"
                joiner($uno, $due, $tre)
            EndIf

        Next

        FileMove($tre, $destfolder & "\" & $nomefinale)
        If FileExists($destfolder & "\PDFjoint_work*.pdf") Then
            FileDelete($destfolder & "\PDFjoint_work*.pdf")
        EndIf

        MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "PDFjoint", "Created new PDF " & @CRLF & @CRLF & $destfolder & "\" & $nomefinale, 2)
    Else
        MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "PDF Joint", "No files selected to Join !")

    EndIf
EndFunc   ;==>PREJoiner

Func joiner($uno, $due, $tre)

    Local $cmd1 = "C:\autoit\PDFjoint\gswin64c.exe"

    Local $cmd2 = '-dBATCH -dNOPAUSE -dQUIET -sDEVICE=pdfwrite -sOutputFile=' & '"' & $tre & '"' & ' ' & '"' & $uno & '"' & ' ' & $due

    ShellExecuteWait($cmd1, $cmd2, "", "open", @SW_HIDE)

EndFunc   ;==>joiner


Func PREExtractor()

    Local $prange = InputBox("PDFjoint -   PDF Pages to extract ", "Write from page to page (example 2,3)", ",")

    Local $arange = _ArrayFromString($prange, ",")

    If IsArray($aPDF) Then

        Local $nomeoutfile, $nomeinfile
        For $i = 1 To $aPDF[0][0]

            $nomeinfile = filenamesanitizer($aPDF[$i][0] & $aPDF[$i][1] & $aPDF[$i][2])

            $nomeoutfile = $destfolder & "\" & onlynamesanitizer($aPDF[$i][1]) & "_pages_" & $arange[0] & "-" & $arange[1] & $aPDF[$i][2]

            Extractor($nomeinfile, $nomeoutfile, $arange[0], $arange[1])

        Next
        MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "PDFjoint", "extracted new PDF(s) in " & @CRLF & @CRLF & $destfolder & "\", 2)
    Else
        MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "PDFJoint", "No files selected to perform extraction !")

    EndIf

EndFunc   ;==>PREExtractor

Func Extractor($in, $out, $dapag, $apag)

    Local $cmd1 = "C:\autoit\PDFjoint\gswin64c.exe"

    Local $cmd2 = '-sDEVICE=pdfwrite -dNOPAUSE -dBATCH  -dFirstPage=' & '"' & $dapag & '"' & ' -dLastPage=' & '"' & $apag & '"' & ' -sOutputFile=' & $out & ' ' & $in

    ShellExecuteWait($cmd1, $cmd2, "", "open", @SW_HIDE)

EndFunc   ;==>Extractor

Func filenamesanitizer($filepathname) ; accept a complete path  + filename and sanitize the name plus move in a temp folder
    Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""

    Local $aPathSplit = _PathSplit($filepathname, $sDrive, $sDir, $sFileName, $sExtension)

    Local $nomefile = StringStripCR($aPathSplit[3])
    $nomefile = StringStripWS($nomefile, 8)
    $nomefile = StringReplace($nomefile, ".", "")
    $nomefile = "c:\temp\" & $nomefile & $aPathSplit[4]

    If Not FileExists("c:\temp") Then DirCreate("c:\temp")
    FileCopy($filepathname, $nomefile, 1)

    Return $nomefile
EndFunc   ;==>filenamesanitizer

Func onlynamesanitizer($onlyname)
    $onlyname = StringStripWS($onlyname, 8)
    $onlyname = StringReplace($onlyname, ".", "")

    Return $onlyname

EndFunc   ;==>onlynamesanitizer

#EndRegion PDF JOIner