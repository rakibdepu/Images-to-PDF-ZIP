#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
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
#include <GuiListBox.au3>
#include <Array.au3>
#include <GuiListView.au3>

Opt("MustDeclareVars", 1) ;0=no, 1=require pre-declaration
Opt("TrayAutoPause", 0) ;0=no pause, 1=Pause
Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Global $trayShow = TrayCreateItem("Show UI")
TrayCreateItem("") ; Create a separator line.
Global $trayExit = TrayCreateItem("Exit")
TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

Global Const $sGUI_Show_Title = 0
Global Const $sAppName = "- Drag and drop -"
Global Const $sLabel_Title = "Drag and drop files and folders HERE !"
Global Const $sLabel_Task = "Or click the button to browse."
Global Const $sLabel_Status = "READY !"

#Region ### START GUI section ###
Global $hGUI
If $sGUI_Show_Title Then
	$hGUI = GUICreate($sAppName, 600, 280, -1, -1, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
Else
	$hGUI = GUICreate($sAppName, 600, 280, -1, -1, BitOR($WS_POPUP, $WS_BORDER), BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
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
Global $idButton_SelectFolder = GUICtrlCreateButton("Select fol&der", 472, 43, 99, 30, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idButton_Close = GUICtrlCreateButton("x", 579, 8, 17, 17, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Exit")
Global $idButton_About = GUICtrlCreateButton("A", 579, 34, 17, 17, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Show About")
Global $idButton_Minimizes = GUICtrlCreateButton("-", 579, 58, 17, 17, BitOR($BS_CENTER, $BS_VCENTER, $BS_FLAT))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetTip(-1, "Minimizes Windows")
Global $idProgress_Total = GUICtrlCreateProgress(1, 1, 596, 4)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idProgress_Current = GUICtrlCreateProgress(1, 75, 596, 4)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $idListFiles = GUICtrlCreateList("", 0, 81, 600, 200)
Global $cDrop_Dummy = GUICtrlCreateDummy() ;Dummy control recieves notifications on filedrop
Global $hGUI_AccelTable[6][2] = [["^1", $idButton_About], ["^2", $idButton_BrowseFiles], ["^3", $idButton_SelectFolder], ["^4", $idButton_Close], ["^5", $idIcon], ["^6", $idButton_Minimizes]]
GUISetAccelerators($hGUI_AccelTable, $hGUI)
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
GUIRegisterMsg($WM_DROPFILES, "_WM_DROPFILES")
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
				_BrowseFolders()
				_GUI_OnStandby()
			Case $idButton_BrowseFiles
				_GUI_OnProgress()
				_BrowseFiles()
				_GUI_OnStandby()
			Case $cDrop_Dummy
				_GUI_OnProgress()
				_On_Drop(GUICtrlRead($cDrop_Dummy))
				_GUI_OnStandby()
		EndSwitch
		;_TRAY_SwitchMsg()
	WEnd
EndIf
; * -----:|
Func _BrowseFiles()
	Local $zListFileIN, $zFileIN = FileOpenDialog("Select Files", @WorkingDir, "All File (*)", $FD_FILEMUSTEXIST + $FD_MULTISELECT, "", $hGUI)             ;1+4
	If Not @error Then
		If StringInStr($zFileIN, "|") Then
			$zListFileIN = StringSplit($zFileIN, "|")
			For $i = 2 To $zListFileIN[0]
				GUICtrlSetData($idListFiles, $zListFileIN[1] & "\" & $zListFileIN[$i])
			Next
		Else
			GUICtrlSetData($idListFiles, $zFileIN)
		EndIf
	EndIf
	ListUpdate()
EndFunc   ;==>_BrowseFiles
Func _BrowseFolders()
	Local $zFolderIN = FileSelectFolder("Select a folder", "", Default, Default, $hGUI)
	If Not @error Then
		Local $aList[1000][2] = [[0]] ;[FileName][FileSz]
		Find_AllFiles($zFolderIN, $aList, 100)             ;Recursively finds files
		For $i = 1 To $aList[0][0]             ;Dumps found files to listview
			GUICtrlSetData($idListFiles, $aList[$i][0])
		Next
	EndIf
	ListUpdate()
EndFunc   ;==>_BrowseFolders
Func _On_Drop($hDrop)
	Local $aDrop_List = _WinAPI_DragQueryFileEx($hDrop, 0) ; 0 = Returns files and folders
	Local $aList[1000][2] = [[0]] ;[FileName][FileSz]
	For $i = 1 To $aDrop_List[0]
		GUICtrlSetData($idListFiles, $aDrop_List[$i]) ;Dumps dropped files to listview
		If StringLen($aDrop_List[$i]) < 4 Then MsgBox(0, "Test", "This Will Take a While Message Etc...")
		Find_AllFiles($aDrop_List[$i], $aList, 100) ;Recursively finds files
	Next
	_GUICtrlListBox_BeginUpdate($idListFiles)
	For $i = 1 To $aList[0][0] ;Dumps found files to listview
		GUICtrlSetData($idListFiles, $aList[$i][0])
	Next
	_GUICtrlListBox_EndUpdate($idListFiles)
	ListUpdate()
EndFunc   ;==>_On_Drop
Func ListUpdate()
	Local $Count, $aListToArray, $aArrayToInput, $aArrayToTrim
	$Count = _GUICtrlListBox_GetCount($idListFiles)
	$aListToArray = ""
	For $i = 0 to $Count - 1
		$aListToArray &= _GUICtrlListBox_GetText($idListFiles, $i) & "|"
	Next
	$aArrayToTrim = StringTrimRight($aListToArray, 1)
	If StringInStr($aArrayToTrim, "|") Then
		$aArrayToInput = StringSplit($aArrayToTrim, "|")
		If IsArray($aArrayToInput) Then
			For $i = 1 To $aArrayToInput[0]
				_Main_Processing($aArrayToInput[$i], $i, $aArrayToInput[0])
			Next
		EndIf
	Else
		_Main_Processing($aArrayToTrim, 1, 1)
	EndIf
EndFunc
Func _Main_Processing($sFilePath, $nCurrent = 0, $nTotal = 0)
	;_GUI_SwitchMsg()
	;_TRAY_SwitchMsg()
	$sPercent = Round(($nCurrent / $nTotal) * 100, 2)
	GUICtrlSetData($idProgress_Total, $sPercent)
	ConsoleWrite("- Percent: " & $sPercent & " %" & @CRLF)
	GUICtrlSetData($idProgress_Current, 0)
	GUICtrlSetData($idLabel_Titles, "Processing " & $nCurrent & "/" & $nTotal & " folder/files ! ")

	Local $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt
	Local $aPathSplit = _SplitPath($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt)
	;Local $sCurrentDirPath= $sDrive&$sCurrentDir;StringRegExpReplace($aPathSplit, '\\[^\\]*$', '')
	;Local $sCurrentDirName =StringRegExpReplace(_PathRemoveBackslash($sCurrentDirPath), '.*\\', '')
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
	If _IsFile($sFilePath) Then
		ConsoleWrite("- Processing file: " & $sFileName & @CRLF)
		GUICtrlSetData($idLabel_Task, "Currently File: " & $sFileName)
		; Your file handler is here!
	Else
		If ($sParentDir == "\" And $sCurrentDir == "") Then ; Is Root Drive
			; Your drive handler is here!
			ConsoleWrite("- Processing drive: " & $sDrive & @CRLF)
			GUICtrlSetData($idLabel_Task, "Currently Drive: " & $sDrive)
		Else
			; Your directory handler is here!
			ConsoleWrite("- Processing directory: " & _PathRemove_Backslash($sPathCurrentDir) & @CRLF)
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
			MsgBox(64, $sAppName & " : About", @YEAR & " © Ðào Văn Trong - Trong.LIVE", Default, $hGUI)
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
EndFunc   ;==>_GUI_OnStandby
; React to items dropped on the GUI
Func _WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	GUICtrlSendToDummy($cDrop_Dummy, $wParam) ;Send the wParam data to the dummy control
EndFunc   ;==>_WM_DROPFILES
; * -----:|
Func _IsFile($sPath)
	If (Not FileExists($sPath)) Then Return SetError(-1, 0, 0)
	If StringInStr(FileGetAttrib($sPath), 'D') <> 0 Then
		Return SetError(0, 0, 0)
	Else
		Return SetError(0, 0, 1)
	EndIf
EndFunc   ;==>_IsFile
Func _DelIt($sPath, $Fc = 1)
	If (Not FileExists($sPath)) Then Return SetError(-1, 0, 1)
	If _IsFile($sPath) Then
		Return SetError(0, 0, _DelFile($sPath, $Fc))
	Else
		Return SetError(0, 0, _RemoveDir($sPath, $Fc))
	EndIf
EndFunc   ;==>_DelIt
Func _DelFile($sPath, $Fc = 1)
	If (Not _IsFile($sPath)) Then
		Return SetError(-1, 0, 0)
	Else
		ConsoleWrite("> DelFile: " & $sPath & @CRLF)
		FileSetAttrib($sPath, "-RSH")
		FileDelete($sPath)
		If $Fc Then
			If FileExists($sPath) Then _TakeOwnership($sPath, "Everyone", $Fc)
			If FileExists($sPath) Then FileDelete($sPath)
			If FileExists($sPath) Then RunWait(@ComSpec & ' /c Del /f /q "' & $sPath & '"', '', @SW_HIDE)
		EndIf
		If FileExists($sPath) Then Return SetError(1, 0, 0)
		Return SetError(0, 0, 1)
	EndIf
EndFunc   ;==>_DelFile
Func _RemoveDir($sPath, $Fc = 1)
	If _IsFile($sPath) Then
		Return SetError(-1, 0, 0)
	Else
		ConsoleWrite("> _RemoveDir: " & $sPath & @CRLF)
		DirRemove($sPath, $Fc)
		If FileExists($sPath) Then _TakeOwnership($sPath, "Everyone", $Fc)
		DirRemove($sPath, $Fc)
		If FileExists($sPath) Then RunWait(@ComSpec & ' /c rmdir "' & $sPath & '" /s /q ', '', @SW_HIDE)
		If FileExists($sPath) Then Return SetError(1, 0, 0)
		Return SetError(0, 0, 1)
	EndIf
EndFunc   ;==>_RemoveDir
Func _Directory_Is_Accessible($sPath, $iTouch = 0)
	If Not FileExists($sPath) Then DirCreate($sPath)
	If Not StringInStr(FileGetAttrib($sPath), "D", 2) Then Return SetError(1, 0, 0)
	Local $iEnum = 0, $maxEnum = 9999, $iRandom = Random(88888888, 99999999, 1)
	If $iTouch Then _TakeOwnership($sPath, "Everyone", 1)
	While FileExists($sPath & "\_IsWritable-" & $iEnum & "-" & $iRandom)
		$iEnum += 1
		If ($iEnum > $maxEnum) Then Return SetError(2, 0, 0)
	WEnd
	Local $iSuccess = DirCreate($sPath & "\_IsWritable-" & $iEnum & "-" & $iRandom)
	Switch $iSuccess
		Case 1
			$iTouch = DirRemove($sPath & "\_IsWritable-" & $iEnum & "-" & $iRandom, 1)
			Return SetError($iTouch < 1, 0, $iTouch)
		Case Else
			Return SetError(3, 0, 0)
	EndSwitch
EndFunc   ;==>_Directory_Is_Accessible
Func _File_Is_Accessible($sFile, $iTouch = 0)
	If ((Not FileExists("\\?\" & $sFile)) Or StringInStr(FileGetAttrib("\\?\" & $sFile), "D", 2)) Then Return SetError(1, 0, 0)
	Local $oFileAttrib = FileGetAttrib("\\?\" & $sFile)
	If $iTouch Then
		FileSetAttrib("\\?\" & $sFile, "-RHS")
		_TakeOwnership($sFile, "Everyone", 1)
	EndIf
	If StringInStr(FileGetAttrib("\\?\" & $sFile), "R", 2) Then Return 3
	Local $hFile = __WinAPI_CreateFileEx("\\?\" & $sFile, 3, 0x80000000 + 0x40000000, 0x00000001 + 0x00000002 + 0x00000004, 0x02000000)
	Local $iReturn = $hFile
	__WinAPI_CloseHandle($hFile)
	If ($iReturn = 0) Then Return 2
	Return 1
EndFunc   ;==>_File_Is_Accessible
Func _TakeOwnership($sFile, $iUserName = "Everyone", $sRecurse = 1)
	If ($iUserName = Default) Or (StringStripWS($iUserName, 8) = '') Then $iUserName = "Everyone"
	If ($sRecurse = Default) Or ($sRecurse = True) Or ($sRecurse > 0) Then
		$sRecurse = 1
	Else
		$sRecurse = 0
	EndIf
	Local $osNotIsEnglish = True
	Switch @OSLang
		Case "0009", "0409", "0809", "0C09", "1009", "1409", "1809", "1C09", "2009", "2409", "2809", "2C09", "3009", "3409", "3C09", "4009", "4409", "4809", "4C09"
			$osNotIsEnglish = False
		Case "3809", "5009", "5409", "5809", "5C09", "6009", "6409"
			$osNotIsEnglish = False
	EndSwitch
	If StringInStr($iUserName, ' ') Then $iUserName = '"' & $iUserName & '"'
	If ($sRecurse = Default) Then $sRecurse = 1
	If Not FileExists($sFile) Then Return SetError(1, 0, $sFile)
	If StringInStr(FileGetAttrib($sFile), 'D') <> 0 Then
		If $sRecurse Then
			RunWait(@ComSpec & ' /c takeown /f "' & $sFile & '" /R /D Y', '', @SW_HIDE)
			If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /T /C /G Administrators:F', '', @SW_HIDE)
			If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /T /C /G Users:F', '', @SW_HIDE)
			RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /T /C /G ' & $iUserName & ':F', '', @SW_HIDE)
			If $osNotIsEnglish Then
				If ($iUserName = "Everyone") Then $iUserName = '*S-1-1-0'
				If ($iUserName = '"' & 'Authenticated Users' & '"') Then $iUserName = '*S-1-5-11'
			EndIf
			If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Administrators:F /T /C /Q', '', @SW_HIDE)
			If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Users:F /T /C /Q', '', @SW_HIDE)
			RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant ' & $iUserName & ':F /T /C /Q', '', @SW_HIDE)
			Return SetError(0, 0, FileSetAttrib($sFile, "-RHS", 1))
		Else
			RunWait(@ComSpec & ' /c takeown /f "' & $sFile & '" ', '', @SW_HIDE)
			If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G Administrators:F', '', @SW_HIDE)
			If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G Users:F', '', @SW_HIDE)
			RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G ' & $iUserName & ':F', '', @SW_HIDE)
			If $osNotIsEnglish Then
				If ($iUserName = "Everyone") Then $iUserName = '*S-1-1-0'
				If ($iUserName = '"' & 'Authenticated Users' & '"') Then $iUserName = '*S-1-5-11'
			EndIf
			If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Administrators:F /C /Q', '', @SW_HIDE)
			If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Users:F /C /Q', '', @SW_HIDE)
			RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant ' & $iUserName & ':F /C /Q', '', @SW_HIDE)
			Return SetError(0, 0, FileSetAttrib($sFile, "-RHS", 0))
		EndIf
	Else
		RunWait(@ComSpec & ' /c takeown /f "' & $sFile & '"', '', @SW_HIDE)
		If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G Administrators:F', '', @SW_HIDE)
		If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G Users:F', '', @SW_HIDE)
		RunWait(@ComSpec & ' /c Echo y|Cacls "' & $sFile & '" /C /G ' & $iUserName & ':F', '', @SW_HIDE)
		If $osNotIsEnglish Then
			If ($iUserName = "Everyone") Then $iUserName = '*S-1-1-0'
			If ($iUserName = '"' & 'Authenticated Users' & '"') Then $iUserName = '*S-1-5-11'
		EndIf
		If $iUserName <> 'Administrators' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Administrators:F /Q', '', @SW_HIDE)
		If $iUserName <> 'Users' Then RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant Users:F /Q', '', @SW_HIDE)
		RunWait(@ComSpec & ' /c iCacls "' & $sFile & '" /grant ' & $iUserName & ':F /Q', '', @SW_HIDE)
		Return SetError(0, 0, FileSetAttrib($sFile, "-RHS"))
	EndIf
	Return $sFile
EndFunc   ;==>_TakeOwnership
; * -----:|
Func _PathGet_Part($sFilePath, $returnPart = 0)
	ConsoleWrite(@CRLF & ";~ + [" & $returnPart & "] PATH IN : " & $sFilePath & @CRLF)
	$sFilePath = _PathFix($sFilePath)
	Local $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt
	Local $aPathSplit = _SplitPath($sFilePath, $sDrive, $sParentDir, $sCurrentDir, $sFileNameNoExt, $sExtension, $sFileName, $sPathParentDir, $sPathCurrentDir, $sPathFileNameNoExt)
	;Local $sCurrentDirPath= $sDrive&$sCurrentDir;StringRegExpReplace($aPathSplit, '\\[^\\]*$', '')
	;Local $sCurrentDirName =StringRegExpReplace(_PathRemoveBackslash($sCurrentDirPath), '.*\\', '')
	ConsoleWrite(";~ - [1] Drive: " & $sDrive & @CRLF)
	ConsoleWrite(";~ - [2] ParentDir: " & $sParentDir & @CRLF)
	ConsoleWrite(";~ - [3] CurrentDir: " & $sCurrentDir & @CRLF)
	ConsoleWrite(";~ - [4] FileName NoExt: " & $sFileNameNoExt & @CRLF)
	ConsoleWrite(";~ - [5] Extension: " & $sExtension & @CRLF)
	ConsoleWrite(";~ - [6] FileName: " & $sFileName & @CRLF)
	ConsoleWrite(";~ - [7] PathParentDir: " & $sPathParentDir & @CRLF)
	ConsoleWrite(";~ - [8] PathCurrentDir: " & $sPathCurrentDir & @CRLF)
	ConsoleWrite(";~ - [9] PathFileName NoExt: " & $sPathFileNameNoExt & @CRLF)
	If (StringStripWS($sFileName, 8) == '') Then ConsoleWrite(";~ ! This path does not contain filenames and extensions!" & @CRLF)
	Switch $returnPart
		Case 1
			;ConsoleWrite(";~ - [1] Drive: " & $sDrive & @CRLF)
			Return $sDrive
		Case 2
			;ConsoleWrite(";~ - [2] ParentDir: " & $sParentDir & @CRLF)
			Return $sParentDir
		Case 3
			;ConsoleWrite(";~ - [3] CurrentDir: " & $sCurrentDir & @CRLF)
			Return $sCurrentDir
		Case 4
			;ConsoleWrite(";~ - [4] FileName NoExt: " & $sFileNameNoExt & @CRLF)
			Return $sFileNameNoExt
		Case 5
			;ConsoleWrite(";~ - [5] Extension: " & $sExtension & @CRLF)
			Return $sExtension
		Case 6
			;ConsoleWrite(";~ - [6] FileName: " & $sFileNameNoExt & $sExtension & @CRLF)
			Return $sFileName
		Case 7
			;ConsoleWrite(";~ - [7] PathParentDir: " & $sDrive & $sParentDir & @CRLF)
			Return $sPathParentDir
		Case 8
			;ConsoleWrite(";~ - [8] PathCurrentDir: " & $sDrive & $sParentDir & $sCurrentDir & @CRLF)
			Return $sPathCurrentDir
		Case 9
			;ConsoleWrite(";~ - [9] PathFileName NoExt: " & $sDrive & $sParentDir & $sCurrentDir & $sFileNameNoExt & @CRLF)
			Return $sPathFileNameNoExt
		Case Else
			ConsoleWrite("! [" & $returnPart & "] PATH OUT : " & $sFilePath & @CRLF)
			Return $sFilePath
	EndSwitch
EndFunc   ;==>_PathGet_Part
; * -----:|
Func _SplitPath($sFilePath, ByRef $sDrive, ByRef $sParentDir, ByRef $sCurrentDir, ByRef $sFileNameNoExt, ByRef $sExtension, ByRef $sFileName, ByRef $sPathParentDir, ByRef $sPathCurrentDir, ByRef $sPathFileNameNoExt)
	;ConsoleWrite(@CRLF & ";~ + PATH IN : " & $sFilePath & @CRLF)
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
Func Find_AllFiles($sPath, ByRef $aList, $iMaxRecursion, $bRecurse = False)

	Local $tData = DllStructCreate($tagWIN32_FIND_DATA)

	Local $sFile
	If StringRight($sPath, 1) <> "\" Then $sPath &= "\"
	Local $hSearch = _WinAPI_FindFirstFile($sPath & '*', $tData)
	While Not @error
		$sFile = DllStructGetData($tData, 'cFileName')
		Switch $sFile
			Case '.', '..'

			Case Else
				If Not BitAND(DllStructGetData($tData, 'dwFileAttributes'), $FILE_ATTRIBUTE_DIRECTORY) Then
					$aList[0][0] += 1
					If $aList[0][0] > UBound($aList) - 1 Then
						ReDim $aList[UBound($aList) + 4096][2]
					EndIf
					$aList[$aList[0][0]][0] = $sPath & $sFile ;;$sFile ;if you want full path.. $sPath & "\" & $sFile
					$aList[$aList[0][0]][1] = _WinAPI_MakeQWord(DllStructGetData($tData, 'nFileSizeLow'), DllStructGetData($tData, 'nFileSizeHigh'))
				ElseIf $bRecurse Then
					If $iMaxRecursion > 0 Then
						Find_AllFiles($sPath & $sFile, $aList, $iMaxRecursion - 1, $bRecurse)
					Else
						MsgBox(0, "Error", "Max Recursion Exceeded")
					EndIf
				EndIf
		EndSwitch
		_WinAPI_FindNextFile($hSearch, $tData)
	WEnd
	_WinAPI_FindClose($hSearch)
EndFunc   ;==>Find_AllFiles