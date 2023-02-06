#include <GUIConstantsEx.au3> ; $GUI_EVENT_CLOSE
#include <WindowsConstants.au3>
#include <GuiListBox.au3>

#include <WinAPISys.au3> ;WinAPI_ChangeWindowMessageFilterEx ;_WinAPI_DragQueryFileEx
#include <WinAPIFiles.au3> ;_WinAPI_FindFirstFile

$hGUI = GUICreate("FileDrop", 550, 300, Default, Default, Default, $WS_EX_ACCEPTFILES)

$cList = GUICtrlCreateList("", 0, 0, 450, 300)

$bClear = GUICtrlCreateButton("Clear", 460, 0, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

$cDrop_Dummy = GUICtrlCreateDummy() ;Dummy control recieves notifications on filedrop
GUISetState(@SW_SHOW)

If IsAdmin() Then ; Allow WM_DROPFILES to be received from lower privileged processes (Windows Vista or later)
    _WinAPI_ChangeWindowMessageFilterEx($hGUI, $WM_COPYGLOBALDATA, $MSGFLT_ALLOW)
    _WinAPI_ChangeWindowMessageFilterEx($hGUI, $WM_DROPFILES, $MSGFLT_ALLOW)
EndIf

; Register $WM_DROPFILES function to detect drops anywhere on the GUI
GUIRegisterMsg($WM_DROPFILES, "_WM_DROPFILES")

While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit
        Case $cDrop_Dummy
            _On_Drop(GUICtrlRead($cDrop_Dummy))
        Case $bClear
            GUICtrlSetData($cList, "")
    EndSwitch
WEnd

Func _On_Drop($hDrop)
    Local $aDrop_List = _WinAPI_DragQueryFileEx($hDrop, 0) ; 0 = Returns files and folders
    Local $aList[1000][2] = [[0]] ;[FileName][FileSz]

    For $i = 1 To $aDrop_List[0]
        GUICtrlSetData($cList, $aDrop_List[$i]) ;Dumps dropped files to listview
        IF StringLen($aDrop_List[$i]) < 4 Then MsgBox(0,"Test", "This Will Take a While Message Etc...")
        Find_AllFiles($aDrop_List[$i], $aList, 100) ;Recursively finds files
    Next
    _GUICtrlListBox_BeginUpdate ($cList)
    For $i = 1 To $aList[0][0] ;Dumps found files to listview
        GUICtrlSetData($cList, $aList[$i][0])
    Next
    _GUICtrlListBox_EndUpdate ($cList)
EndFunc   ;==>_On_Drop


; React to items dropped on the GUI
Func _WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $lParam
    GUICtrlSendToDummy($cDrop_Dummy, $wParam) ;Send the wParam data to the dummy control
EndFunc   ;==>_WM_DROPFILES

Func Find_AllFiles($sPath, ByRef $aList, $iMaxRecursion, $bRecurse = True)

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
                    $aList[$aList[0][0]][0] = $sPath & $sFile;;$sFile ;if you want full path.. $sPath & "\" & $sFile
                    $aList[$aList[0][0]][1] = _WinAPI_MakeQWord(DllStructGetData($tData, 'nFileSizeLow'), DllStructGetData($tData, 'nFileSizeHigh'))
                Elseif $bRecurse Then
                    if $iMaxRecursion > 0 Then
                    Find_AllFiles($sPath & $sFile, $aList, $iMaxRecursion - 1, $bRecurse)
                    Else
                        Msgbox(0,"Error", "Max Recursion Exceeded")
                    EndIF
                EndIf
        EndSwitch
        _WinAPI_FindNextFile($hSearch, $tData)
    WEnd
    _WinAPI_FindClose($hSearch)
EndFunc   ;==>Find_AllFiles