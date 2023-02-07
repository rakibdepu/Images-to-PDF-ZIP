#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=faviconO.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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

Opt("MustDeclareVars", 1) ;0=no, 1=require pre-declaration
Opt("TrayAutoPause", 0) ;0=no pause, 1=Pause
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Global $trayShow = TrayCreateItem("Show UI")
TrayCreateItem("") ; Create a separator line.
Global $trayExit = TrayCreateItem("Exit")
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

Global Const $sGUI_Show_Title = 0
Global Const $sAppName = "- Images to PDF & ZIP -"
Global Const $sLabel_Title = "Drag and drop files and folders"
Global Const $sLabel_Task = "Or click the button to browse."
Global Const $sLabel_Status = "READY !"

#Region ### START GUI section ###
Local $SplashScreenGui = GUICreate("SplashScreen", 600, 360, -1,-1,$WS_POPUP)
WinSetTrans($SplashScreenGui, "", 170)
Local $Pic1 = GUICtrlCreatePic("splash.gif", 0, 0, 600, 360)
GUISetState(@SW_SHOW,$SplashScreenGui)
Sleep(5000)
GUISetState(@SW_HIDE,$SplashScreenGui)
Global $hGUI
If $sGUI_Show_Title Then
    $hGUI = GUICreate($sAppName, 400, 200, -1, -1, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
Else
    $hGUI = GUICreate($sAppName, 400, 200, -1, -1, BitOR($WS_POPUP, $WS_BORDER), BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
EndIf
_WinAPI_DwmSetWindowAttribute($hGUI, 33, 2)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idIcon = GUICtrlCreateIcon(@WorkingDir & "\favicon.ico", -1, 5, 50, 113, 68, BitOR($GUI_SS_DEFAULT_ICON, $SS_CENTERIMAGE))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Set/UnSet Windows on TOP")
Global $idLabel_BG = GUICtrlCreateLabel($sAppName, 0, 5, 400, 32, BitOR($SS_CENTER, $SS_CENTERIMAGE, $SS_NOPREFIX), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetFont(-1, 18, 600, 4, "Grab Community EN v2.0 Inline", 5)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idProgress_Total = GUICtrlCreateProgress(0, 42, 400, 5)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUIStartGroup()
Global $idButton_pdf = GUICtrlCreateRadio("PDF", 129, 52, 40, 18, BitOR($BS_AUTORADIOBUTTON, $BS_VCENTER), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_both = GUICtrlCreateRadio("Both", 129, 75, 40, 18, BitOR($BS_AUTORADIOBUTTON, $BS_VCENTER), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_zip = GUICtrlCreateRadio("ZIP", 129, 98, 40, 18, BitOR($BS_AUTORADIOBUTTON, $BS_VCENTER), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idInput_term = GUICtrlCreateInput("", 179, 52, 40, 18)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUIStartGroup()
Global $idButton_Incl = GUICtrlCreateRadio("Incl.", 179, 75, 40, 18, BitOR($BS_AUTORADIOBUTTON, $BS_VCENTER), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_Excl = GUICtrlCreateRadio("Excl.", 179, 98, 40, 18, BitOR($BS_AUTORADIOBUTTON, $BS_VCENTER), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_BrowseFiles = GUICtrlCreateButton("Browse" & @CRLF & "&files", 229, 52, 64, 64, BitOR($BS_CENTER, $BS_VCENTER, $BS_MULTILINE, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idButton_SelectFolder = GUICtrlCreateButton("Select" & @CRLF & "fol&der", 303, 52, 64, 64, BitOR($BS_CENTER, $BS_VCENTER, $BS_MULTILINE, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idButton_Close = GUICtrlCreateButton("X", 377, 52, 18, 18, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Exit")
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idButton_About = GUICtrlCreateButton("A", 377, 75, 18, 18, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Show About")
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idButton_Minimizes = GUICtrlCreateButton("-", 377, 98, 18, 18, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Minimizes Windows")
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idProgress_Current = GUICtrlCreateProgress(0, 121, 400, 5)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idLabel_Titles = GUICtrlCreateLabel($sLabel_Title, 0, 131, 400, 18, BitOR($SS_CENTER, $SS_CENTERIMAGE), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idLabel_Task = GUICtrlCreateLabel($sLabel_Task, 0, 154, 400, 18, BitOR($SS_CENTER, $SS_CENTERIMAGE), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $idLabel_Status = GUICtrlCreateLabel($sLabel_Status, 0, 177, 400, 18, BitOR($SS_CENTER, $SS_CENTERIMAGE), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8.5, 400, 0, "Ubuntu", 5)
Global $hGUI_AccelTable[7][2] = [["^1", $idButton_About], ["^2", $idButton_BrowseFiles], ["^3", $idButton_SelectFolder], ["^4", $idButton_Close], ["^5", $idIcon], ["^6", $idButton_Minimizes], ["^7", $idInput_term]]
GUISetAccelerators($hGUI_AccelTable, $hGUI)

Global $hGUI_Child = GUICreate("About", 400, 200, -1, -1, $WS_POPUP, $WS_EX_TOPMOST, $hGUI)
;GUISetBkColor(0x000000, $hGUI_Child)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

Local $idLabelAbout1 = GUICtrlCreateLabel("The author disclaims copyright to this source code." & @CRLF & "In place of a legal notice, here is a blessing:", 10, 40, 380, 45, BitOR($SS_CENTER, $BS_MULTILINE))
;GUICtrlSetColor($idLabelAbout1, 0xFFFFFF)
GUICtrlSetFont(-1, 11, 400, 4, "Ubuntu", 5)
Local $idLabelAbout2 = GUICtrlCreateLabel("May you do good and not evil." & @CRLF & "May you find forgiveness for yourself and forgive others." & @CRLF & "May you share freely, never taking more than you give.", 25, 85, 350, 120)
;GUICtrlSetColor($idLabelAbout2, 0xFFFFFF)
GUICtrlSetFont(-1, 11, 400, 0, "Ubuntu", 5)
_GuiRoundCorners($hGUI_Child, 0, 0, 7, 7)
WinSetTrans($hGUI_Child, "", 0)
GUISetState(@SW_HIDE, $hGUI_Child)

#EndRegion ### START GUI section ###
HotKeySet('^5', '_GUI_SetOnTop')
Global $GuiOnTop = 0, $onWorking = 0
WinSetTrans($hGUI, "", 91)
_GUI_OnProgress()
_GUI_SHOW()
_GUI_OnStandby()
WinSetTrans($hGUI, "", 200)

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
            Case $idButton_SelectFolder
                _GUI_OnProgress()
                ; Display an open dialog to select a folder.
                Local $zFolderIN = FileSelectFolder("Select a folder", "", Default, Default, $hGUI)
                If Not @error Then
                    _Main_Processing($zFolderIN, 1, 1)
                EndIf
                _GUI_OnStandby()
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
    ConsoleWrite("Total Progress: " & $sPercent & "%" & @CRLF)
    GUICtrlSetData($idProgress_Current, 0)
    GUICtrlSetData($idLabel_Titles, "Total selected " & $nTotal & " files. (" & $sPercent & "%)")

    Local $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt
    Local $aPathSplit = _SplitPath($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt)
    ;Local $sCurrentDirPath= $sDrive&$sCurrentDir;StringRegExpReplace($aPathSplit, '\\[^\\]*$', '')
    ;Local $sCurrentDirName =StringRegExpReplace(_PathRemoveBackslash($sCurrentDirPath), '.*\\', '')
    ConsoleWrite("[1] Drive: " & $sDrive & @CRLF)
    ConsoleWrite("[2] ParentDir: " & $sParentDir & @CRLF)
    ConsoleWrite("[3] CurrentDir: " & $sCurrentDir & @CRLF)
    ConsoleWrite("[4] FileName NoExt: " & $sFileNameNoExt & @CRLF)
    ConsoleWrite("[5] Extension: " & $sExtension & @CRLF)
    ConsoleWrite("[6] FileName: " & $sFileName & @CRLF)
    ConsoleWrite("[7] PathParentDir: " & $sPathParentDir & @CRLF)
    ConsoleWrite("[8] PathCurrentDir: " & $sPathCurrentDir & @CRLF)
    ConsoleWrite("[9] PathFileName NoExt: " & $sPathFileNameNoExt & @CRLF)
    ConsoleWrite("Total selected " & $nTotal & " files. (" & $sPercent & "%)" & @CRLF)
    If _IsFile($sFilePath) Then
        ConsoleWrite("Processing now: (" & $nCurrent & ") " & $sFileName & @CRLF)
        GUICtrlSetData($idLabel_Task, "Processing now: (" & $nCurrent & ") " & $sFileName)
        ; Your file handler is here!
    Else
        If (FileExists($sPathCurrentDir)) Then ; Is Root Drive
            ; Your drive handler is here!
            ConsoleWrite("Processing directory: " & _PathRemove_Backslash($sPathCurrentDir) & @CRLF)
            GUICtrlSetData($idLabel_Task, "Currently Folder: " & _PathRemove_Backslash($sCurrentDir))
        Else
            ; Your directory handler is here!
            ConsoleWrite("Processing drive: " & $sDrive & @CRLF)
            GUICtrlSetData($idLabel_Task, "Currently Drive: " & $sDrive)
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
            ;MsgBox(64, "About", "© " & @YEAR & " Rakibul Hafiz", Default, $hGUI)
			GUISetState(@SW_MINIMIZE, $hGUI)
			GUISetState(@SW_SHOW, $hGUI_Child)
			
			For $i = 0 To 150 Step 8
				WinSetTrans($hGUI_Child, "", $i)
				Sleep(10)
			Next
			Sleep(5000)
			For $i = 150 To 0 Step -4
				WinSetTrans($hGUI_Child, "", $i)
				Sleep(10)
			Next
			
			GUISetState(@SW_HIDE, $hGUI_Child)
			_GUI_SHOW()
			
        Case $idButton_Minimizes
            GUISetState(@SW_MINIMIZE, $hGUI)
        Case $idIcon
            _GUI_SetOnTop()
        Case $GUI_EVENT_CLOSE, $idButton_Close
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
    GUICtrlSetState($idButton_SelectFolder, $GUI_DISABLE)     ; DISABLE
    GUICtrlSetState($idLabel_BG, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idIcon, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Titles, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Task, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idLabel_Status, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_BrowseFiles, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_SelectFolder, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_Close, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_About, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_Minimizes, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idProgress_Total, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idProgress_Current, $GUI_NODROPACCEPTED) ; NODROPACCEPTED
    GUICtrlSetState($idButton_pdf, $GUI_NODROPACCEPTED)
    GUICtrlSetState($idButton_zip, $GUI_NODROPACCEPTED)
    GUICtrlSetState($idButton_both, $GUI_NODROPACCEPTED)
    GUICtrlSetState($idInput_term, $GUI_NODROPACCEPTED)
    GUICtrlSetState($idButton_Incl, $GUI_NODROPACCEPTED)
    GUICtrlSetState($idButton_Excl, $GUI_NODROPACCEPTED)

EndFunc   ;==>_GUI_OnProgress
Func _GUI_OnStandby()
    $onWorking = 0
    GUICtrlSetData($idLabel_Titles, $sLabel_Title)
    GUICtrlSetData($idLabel_Task, $sLabel_Task)
    GUICtrlSetData($idLabel_Status, $sLabel_Status)
    GUICtrlSetData($idProgress_Total, 100)
    GUICtrlSetData($idProgress_Current, 100)
    GUICtrlSetState($idButton_BrowseFiles, $GUI_ENABLE)     ; ENABLE
    GUICtrlSetState($idButton_SelectFolder, $GUI_ENABLE)     ; ENABLE
    GUICtrlSetState($idLabel_BG, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idIcon, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Titles, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Task, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idLabel_Status, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_BrowseFiles, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_SelectFolder, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_Close, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_About, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_Minimizes, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idProgress_Total, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idProgress_Current, $GUI_DROPACCEPTED) ; DROPACCEPTED
    GUICtrlSetState($idButton_pdf, $GUI_DROPACCEPTED)
    GUICtrlSetState($idButton_zip, $GUI_DROPACCEPTED)
    GUICtrlSetState($idButton_both, $GUI_DROPACCEPTED)
    GUICtrlSetState($idInput_term, $GUI_DROPACCEPTED)
    GUICtrlSetState($idButton_Incl, $GUI_DROPACCEPTED)
    GUICtrlSetState($idButton_Excl, $GUI_DROPACCEPTED)
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
Func _SplitPath($sFilePath, ByRef $sDrive, ByRef $sParentDir, ByRef $sCurrentDir, ByRef $sFileNameNoExt, ByRef $sExtension, ByRef $sFileName, ByRef $sPathParentDir, ByRef $sPathCurrentDir, ByRef $sPathFileNameNoExt)
    ;ConsoleWrite(@CRLF & " + PATH IN : " & $sFilePath & @CRLF)
    $sFilePath = _PathFix($sFilePath)
    _PathSplit_Ex($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension)
    ;Local $sCurrentDirPath= $sDrive&$sCurrentDir;StringRegExpReplace($aPathSplit, '\\[^\\]*$', '')
    ;Local $sCurrentDirName =StringRegExpReplace(_PathRemoveBackslash($sCurrentDirPath), '.*\\', '')
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
        ;$sPath = StringRegExpReplace($sPath, "\\$", "")
        ;$sPath = StringRegExpReplace($sPath, "(.*)(\\)\z", "$1")
        $sPath = StringTrimRight($sPath, 1)
    EndIf
    Return $sPath
EndFunc   ;==>_PathRemove_Backslash
Func _GuiRoundCorners($hWnd, $iLeftRect, $iTopRect, $iWidthEllipse, $iHeightEllipse)
    Local $aPos = 0, $aRet = 0

    $aPos = WinGetPos($hWnd)

    $aRet = DllCall("gdi32.dll", "long", "CreateRoundRectRgn", "long", $iLeftRect, "long", $iTopRect, "long", $aPos[2], "long", $aPos[3], "long", $iWidthEllipse, "long", $iHeightEllipse)
    If Not @error And $aRet[0] Then
        $aRet = DllCall("user32.dll", "long", "SetWindowRgn", "hwnd", $hWnd, "long", $aRet[0], "int", 1)
        If Not @error And $aRet[0] Then Return 1
    EndIf

    Return 0
EndFunc   ;==>_GuiRoundCorners
; * -----:|