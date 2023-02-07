#EndRegion Attributes Tool
#NoTrayIcon
#RequireAdmin
;~ Opt("MustDeclareVars", 1)
Opt("GUICloseOnESC", 0)
Opt("TrayAutoPause", 0)
#Region ###include###
;~ #include <APIConstants.au3>
;~ #include <ButtonConstants.au3>
;~ #include <EditConstants.au3>
;~ #include <GUIConstantsEx.au3>
;~ #include <ProgressConstants.au3>
;~ #include <StaticConstants.au3>
;~ #include <WindowsConstants.au3>
;~ #include <File.au3>
;~ #include <WinAPIEx.au3>
;~ #include <WinAPIShPath.au3>
;~ #include <Misc.au3>
Global Const $MSGFLT_ALLOW = 1
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $DT_ALL = "ALL"
Global Const $DS_READY = "READY"
Global Const $FT_MODIFIED = 0
Global Const $FT_CREATED = 1
Global Const $FT_ACCESSED = 2
Global Const $FD_MULTISELECT = 4
Global Const $FLTA_FILESFOLDERS = 0
Global Const $FLTAR_FILESFOLDERS = 0
Global Const $FLTAR_NORECUR = 0
Global Const $FLTAR_NOSORT = 0
Global Const $FLTAR_RELPATH = 1
Global Const $STR_NOCASESENSEBASIC = 2
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Global Const $BS_MULTILINE = 0x2000
Global Const $BS_VCENTER = 0x0C00
Global Const $BS_FLAT = 0x8000
Global Const $GUI_SS_DEFAULT_CHECKBOX = 0
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_EVENT_DROPPED = -13
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_CHECKED = 1
Global Const $GUI_DROPACCEPTED = 8
Global Const $GUI_BKCOLOR_TRANSPARENT = -2
Global Const $GUI_WS_EX_PARENTDRAG = 0x00100000
Global Const $PBS_SMOOTH = 1
Global Const $SS_CENTER = 0x1
Global Const $SS_CENTERIMAGE = 0x0200
Global Const $WS_EX_ACCEPTFILES = 0x00000010
Global Const $WS_EX_TOPMOST = 0x00000008
Global Const $WS_EX_WINDOWEDGE = 0x00000100
Global Const $WM_COPYGLOBALDATA = 0x0049
Global Const $WM_COPYDATA = 0x004A
Global Const $WM_DROPFILES = 0x0233
Func _ArrayConcatenate(ByRef $aArrayTarget, Const ByRef $aArraySource, $iStart = 0)
	If $iStart = Default Then $iStart = 0
	If Not IsArray($aArrayTarget) Then Return SetError(1, 0, -1)
	If Not IsArray($aArraySource) Then Return SetError(2, 0, -1)
	Local $iDim_Total_Tgt = UBound($aArrayTarget, $UBOUND_DIMENSIONS)
	Local $iDim_Total_Src = UBound($aArraySource, $UBOUND_DIMENSIONS)
	Local $iDim_1_Tgt = UBound($aArrayTarget, $UBOUND_ROWS)
	Local $iDim_1_Src = UBound($aArraySource, $UBOUND_ROWS)
	If $iStart < 0 Or $iStart > $iDim_1_Src - 1 Then Return SetError(6, 0, -1)
	Switch $iDim_Total_Tgt
		Case 1
			If $iDim_Total_Src <> 1 Then Return SetError(4, 0, -1)
			ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart]
			For $i = $iStart To $iDim_1_Src - 1
				$aArrayTarget[$iDim_1_Tgt + $i - $iStart] = $aArraySource[$i]
			Next
		Case 2
			If $iDim_Total_Src <> 2 Then Return SetError(4, 0, -1)
			Local $iDim_2_Tgt = UBound($aArrayTarget, $UBOUND_COLUMNS)
			If UBound($aArraySource, $UBOUND_COLUMNS) <> $iDim_2_Tgt Then Return SetError(5, 0, -1)
			ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart][$iDim_2_Tgt]
			For $i = $iStart To $iDim_1_Src - 1
				For $j = 0 To $iDim_2_Tgt - 1
					$aArrayTarget[$iDim_1_Tgt + $i - $iStart][$j] = $aArraySource[$i][$j]
				Next
			Next
		Case Else
			Return SetError(3, 0, -1)
	EndSwitch
	Return UBound($aArrayTarget, $UBOUND_ROWS)
EndFunc   ;==>_ArrayConcatenate
Func _ArrayReverse(ByRef $aArray, $iStart = 0, $iEnd = 0)
	If $iStart = Default Then $iStart = 0
	If $iEnd = Default Then $iEnd = 0
	If Not IsArray($aArray) Then Return SetError(1, 0, 0)
	If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(3, 0, 0)
	If Not UBound($aArray) Then Return SetError(4, 0, 0)
	Local $vTmp, $iUBound = UBound($aArray) - 1
	If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
	If $iStart < 0 Then $iStart = 0
	If $iStart > $iEnd Then Return SetError(2, 0, 0)
	For $i = $iStart To Int(($iStart + $iEnd - 1) / 2)
		$vTmp = $aArray[$i]
		$aArray[$i] = $aArray[$iEnd]
		$aArray[$iEnd] = $vTmp
		$iEnd -= 1
	Next
	Return 1
EndFunc   ;==>_ArrayReverse
Func _ArraySort(ByRef $aArray, $iDescending = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0, $iPivot = 0)
	If $iDescending = Default Then $iDescending = 0
	If $iStart = Default Then $iStart = 0
	If $iEnd = Default Then $iEnd = 0
	If $iSubItem = Default Then $iSubItem = 0
	If $iPivot = Default Then $iPivot = 0
	If Not IsArray($aArray) Then Return SetError(1, 0, 0)
	Local $iUBound = UBound($aArray) - 1
	If $iUBound = -1 Then Return SetError(5, 0, 0)
	If $iEnd = Default Then $iEnd = 0
	If $iEnd < 1 Or $iEnd > $iUBound Or $iEnd = Default Then $iEnd = $iUBound
	If $iStart < 0 Or $iStart = Default Then $iStart = 0
	If $iStart > $iEnd Then Return SetError(2, 0, 0)
	If $iDescending = Default Then $iDescending = 0
	If $iPivot = Default Then $iPivot = 0
	If $iSubItem = Default Then $iSubItem = 0
	Switch UBound($aArray, $UBOUND_DIMENSIONS)
		Case 1
			If $iPivot Then
				__ArrayDualPivotSort($aArray, $iStart, $iEnd)
			Else
				__ArrayQuickSort1D($aArray, $iStart, $iEnd)
			EndIf
			If $iDescending Then _ArrayReverse($aArray, $iStart, $iEnd)
		Case 2
			If $iPivot Then Return SetError(6, 0, 0)
			Local $iSubMax = UBound($aArray, $UBOUND_COLUMNS) - 1
			If $iSubItem > $iSubMax Then Return SetError(3, 0, 0)
			If $iDescending Then
				$iDescending = -1
			Else
				$iDescending = 1
			EndIf
			__ArrayQuickSort2D($aArray, $iDescending, $iStart, $iEnd, $iSubItem, $iSubMax)
		Case Else
			Return SetError(4, 0, 0)
	EndSwitch
	Return 1
EndFunc   ;==>_ArraySort
Func __ArrayQuickSort1D(ByRef $aArray, Const ByRef $iStart, Const ByRef $iEnd)
	If $iEnd <= $iStart Then Return
	Local $vTmp
	If ($iEnd - $iStart) < 15 Then
		Local $vCur
		For $i = $iStart + 1 To $iEnd
			$vTmp = $aArray[$i]
			If IsNumber($vTmp) Then
				For $j = $i - 1 To $iStart Step -1
					$vCur = $aArray[$j]
					If ($vTmp >= $vCur And IsNumber($vCur)) Or (Not IsNumber($vCur) And StringCompare($vTmp, $vCur) >= 0) Then ExitLoop
					$aArray[$j + 1] = $vCur
				Next
			Else
				For $j = $i - 1 To $iStart Step -1
					If (StringCompare($vTmp, $aArray[$j]) >= 0) Then ExitLoop
					$aArray[$j + 1] = $aArray[$j]
				Next
			EndIf
			$aArray[$j + 1] = $vTmp
		Next
		Return
	EndIf
	Local $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)], $bNum = IsNumber($vPivot)
	Do
		If $bNum Then
			While ($aArray[$L] < $vPivot And IsNumber($aArray[$L])) Or (Not IsNumber($aArray[$L]) And StringCompare($aArray[$L], $vPivot) < 0)
				$L += 1
			WEnd
			While ($aArray[$R] > $vPivot And IsNumber($aArray[$R])) Or (Not IsNumber($aArray[$R]) And StringCompare($aArray[$R], $vPivot) > 0)
				$R -= 1
			WEnd
		Else
			While (StringCompare($aArray[$L], $vPivot) < 0)
				$L += 1
			WEnd
			While (StringCompare($aArray[$R], $vPivot) > 0)
				$R -= 1
			WEnd
		EndIf
		If $L <= $R Then
			$vTmp = $aArray[$L]
			$aArray[$L] = $aArray[$R]
			$aArray[$R] = $vTmp
			$L += 1
			$R -= 1
		EndIf
	Until $L > $R
	__ArrayQuickSort1D($aArray, $iStart, $R)
	__ArrayQuickSort1D($aArray, $L, $iEnd)
EndFunc   ;==>__ArrayQuickSort1D
Func __ArrayQuickSort2D(ByRef $aArray, Const ByRef $iStep, Const ByRef $iStart, Const ByRef $iEnd, Const ByRef $iSubItem, Const ByRef $iSubMax)
	If $iEnd <= $iStart Then Return
	Local $vTmp, $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)][$iSubItem], $bNum = IsNumber($vPivot)
	Do
		If $bNum Then
			While ($iStep * ($aArray[$L][$iSubItem] - $vPivot) < 0 And IsNumber($aArray[$L][$iSubItem])) Or (Not IsNumber($aArray[$L][$iSubItem]) And $iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
				$L += 1
			WEnd
			While ($iStep * ($aArray[$R][$iSubItem] - $vPivot) > 0 And IsNumber($aArray[$R][$iSubItem])) Or (Not IsNumber($aArray[$R][$iSubItem]) And $iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
				$R -= 1
			WEnd
		Else
			While ($iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
				$L += 1
			WEnd
			While ($iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
				$R -= 1
			WEnd
		EndIf
		If $L <= $R Then
			For $i = 0 To $iSubMax
				$vTmp = $aArray[$L][$i]
				$aArray[$L][$i] = $aArray[$R][$i]
				$aArray[$R][$i] = $vTmp
			Next
			$L += 1
			$R -= 1
		EndIf
	Until $L > $R
	__ArrayQuickSort2D($aArray, $iStep, $iStart, $R, $iSubItem, $iSubMax)
	__ArrayQuickSort2D($aArray, $iStep, $L, $iEnd, $iSubItem, $iSubMax)
EndFunc   ;==>__ArrayQuickSort2D
Func __ArrayDualPivotSort(ByRef $aArray, $iPivot_Left, $iPivot_Right, $bLeftMost = True)
	If $iPivot_Left > $iPivot_Right Then Return
	Local $iLength = $iPivot_Right - $iPivot_Left + 1
	Local $i, $j, $k, $iAi, $iAk, $iA1, $iA2, $iLast
	If $iLength < 45 Then
		If $bLeftMost Then
			$i = $iPivot_Left
			While $i < $iPivot_Right
				$j = $i
				$iAi = $aArray[$i + 1]
				While $iAi < $aArray[$j]
					$aArray[$j + 1] = $aArray[$j]
					$j -= 1
					If $j + 1 = $iPivot_Left Then ExitLoop
				WEnd
				$aArray[$j + 1] = $iAi
				$i += 1
			WEnd
		Else
			While 1
				If $iPivot_Left >= $iPivot_Right Then Return 1
				$iPivot_Left += 1
				If $aArray[$iPivot_Left] < $aArray[$iPivot_Left - 1] Then ExitLoop
			WEnd
			While 1
				$k = $iPivot_Left
				$iPivot_Left += 1
				If $iPivot_Left > $iPivot_Right Then ExitLoop
				$iA1 = $aArray[$k]
				$iA2 = $aArray[$iPivot_Left]
				If $iA1 < $iA2 Then
					$iA2 = $iA1
					$iA1 = $aArray[$iPivot_Left]
				EndIf
				$k -= 1
				While $iA1 < $aArray[$k]
					$aArray[$k + 2] = $aArray[$k]
					$k -= 1
				WEnd
				$aArray[$k + 2] = $iA1
				While $iA2 < $aArray[$k]
					$aArray[$k + 1] = $aArray[$k]
					$k -= 1
				WEnd
				$aArray[$k + 1] = $iA2
				$iPivot_Left += 1
			WEnd
			$iLast = $aArray[$iPivot_Right]
			$iPivot_Right -= 1
			While $iLast < $aArray[$iPivot_Right]
				$aArray[$iPivot_Right + 1] = $aArray[$iPivot_Right]
				$iPivot_Right -= 1
			WEnd
			$aArray[$iPivot_Right + 1] = $iLast
		EndIf
		Return 1
	EndIf
	Local $iSeventh = BitShift($iLength, 3) + BitShift($iLength, 6) + 1
	Local $iE1, $iE2, $iE3, $iE4, $iE5, $t
	$iE3 = Ceiling(($iPivot_Left + $iPivot_Right) / 2)
	$iE2 = $iE3 - $iSeventh
	$iE1 = $iE2 - $iSeventh
	$iE4 = $iE3 + $iSeventh
	$iE5 = $iE4 + $iSeventh
	If $aArray[$iE2] < $aArray[$iE1] Then
		$t = $aArray[$iE2]
		$aArray[$iE2] = $aArray[$iE1]
		$aArray[$iE1] = $t
	EndIf
	If $aArray[$iE3] < $aArray[$iE2] Then
		$t = $aArray[$iE3]
		$aArray[$iE3] = $aArray[$iE2]
		$aArray[$iE2] = $t
		If $t < $aArray[$iE1] Then
			$aArray[$iE2] = $aArray[$iE1]
			$aArray[$iE1] = $t
		EndIf
	EndIf
	If $aArray[$iE4] < $aArray[$iE3] Then
		$t = $aArray[$iE4]
		$aArray[$iE4] = $aArray[$iE3]
		$aArray[$iE3] = $t
		If $t < $aArray[$iE2] Then
			$aArray[$iE3] = $aArray[$iE2]
			$aArray[$iE2] = $t
			If $t < $aArray[$iE1] Then
				$aArray[$iE2] = $aArray[$iE1]
				$aArray[$iE1] = $t
			EndIf
		EndIf
	EndIf
	If $aArray[$iE5] < $aArray[$iE4] Then
		$t = $aArray[$iE5]
		$aArray[$iE5] = $aArray[$iE4]
		$aArray[$iE4] = $t
		If $t < $aArray[$iE3] Then
			$aArray[$iE4] = $aArray[$iE3]
			$aArray[$iE3] = $t
			If $t < $aArray[$iE2] Then
				$aArray[$iE3] = $aArray[$iE2]
				$aArray[$iE2] = $t
				If $t < $aArray[$iE1] Then
					$aArray[$iE2] = $aArray[$iE1]
					$aArray[$iE1] = $t
				EndIf
			EndIf
		EndIf
	EndIf
	Local $iLess = $iPivot_Left
	Local $iGreater = $iPivot_Right
	If (($aArray[$iE1] <> $aArray[$iE2]) And ($aArray[$iE2] <> $aArray[$iE3]) And ($aArray[$iE3] <> $aArray[$iE4]) And ($aArray[$iE4] <> $aArray[$iE5])) Then
		Local $iPivot_1 = $aArray[$iE2]
		Local $iPivot_2 = $aArray[$iE4]
		$aArray[$iE2] = $aArray[$iPivot_Left]
		$aArray[$iE4] = $aArray[$iPivot_Right]
		Do
			$iLess += 1
		Until $aArray[$iLess] >= $iPivot_1
		Do
			$iGreater -= 1
		Until $aArray[$iGreater] <= $iPivot_2
		$k = $iLess
		While $k <= $iGreater
			$iAk = $aArray[$k]
			If $iAk < $iPivot_1 Then
				$aArray[$k] = $aArray[$iLess]
				$aArray[$iLess] = $iAk
				$iLess += 1
			ElseIf $iAk > $iPivot_2 Then
				While $aArray[$iGreater] > $iPivot_2
					$iGreater -= 1
					If $iGreater + 1 = $k Then ExitLoop 2
				WEnd
				If $aArray[$iGreater] < $iPivot_1 Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $aArray[$iGreater]
					$iLess += 1
				Else
					$aArray[$k] = $aArray[$iGreater]
				EndIf
				$aArray[$iGreater] = $iAk
				$iGreater -= 1
			EndIf
			$k += 1
		WEnd
		$aArray[$iPivot_Left] = $aArray[$iLess - 1]
		$aArray[$iLess - 1] = $iPivot_1
		$aArray[$iPivot_Right] = $aArray[$iGreater + 1]
		$aArray[$iGreater + 1] = $iPivot_2
		__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 2, True)
		__ArrayDualPivotSort($aArray, $iGreater + 2, $iPivot_Right, False)
		If ($iLess < $iE1) And ($iE5 < $iGreater) Then
			While $aArray[$iLess] = $iPivot_1
				$iLess += 1
			WEnd
			While $aArray[$iGreater] = $iPivot_2
				$iGreater -= 1
			WEnd
			$k = $iLess
			While $k <= $iGreater
				$iAk = $aArray[$k]
				If $iAk = $iPivot_1 Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $iAk
					$iLess += 1
				ElseIf $iAk = $iPivot_2 Then
					While $aArray[$iGreater] = $iPivot_2
						$iGreater -= 1
						If $iGreater + 1 = $k Then ExitLoop 2
					WEnd
					If $aArray[$iGreater] = $iPivot_1 Then
						$aArray[$k] = $aArray[$iLess]
						$aArray[$iLess] = $iPivot_1
						$iLess += 1
					Else
						$aArray[$k] = $aArray[$iGreater]
					EndIf
					$aArray[$iGreater] = $iAk
					$iGreater -= 1
				EndIf
				$k += 1
			WEnd
		EndIf
		__ArrayDualPivotSort($aArray, $iLess, $iGreater, False)
	Else
		Local $iPivot = $aArray[$iE3]
		$k = $iLess
		While $k <= $iGreater
			If $aArray[$k] = $iPivot Then
				$k += 1
				ContinueLoop
			EndIf
			$iAk = $aArray[$k]
			If $iAk < $iPivot Then
				$aArray[$k] = $aArray[$iLess]
				$aArray[$iLess] = $iAk
				$iLess += 1
			Else
				While $aArray[$iGreater] > $iPivot
					$iGreater -= 1
				WEnd
				If $aArray[$iGreater] < $iPivot Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $aArray[$iGreater]
					$iLess += 1
				Else
					$aArray[$k] = $iPivot
				EndIf
				$aArray[$iGreater] = $iAk
				$iGreater -= 1
			EndIf
			$k += 1
		WEnd
		__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 1, True)
		__ArrayDualPivotSort($aArray, $iGreater + 1, $iPivot_Right, False)
	EndIf
EndFunc   ;==>__ArrayDualPivotSort
Func _FileListToArray($sFilePath, $sFilter = "*", $iFlag = $FLTA_FILESFOLDERS, $bReturnPath = False)
	Local $sDelimiter = "|", $sFileList = "", $sFileName = "", $sFullPath = ""
	$sFilePath = StringRegExpReplace($sFilePath, "[\\/]+$", "") & "\"
	If $iFlag = Default Then $iFlag = $FLTA_FILESFOLDERS
	If $bReturnPath Then $sFullPath = $sFilePath
	If $sFilter = Default Then $sFilter = "*"
	If Not FileExists($sFilePath) Then Return SetError(1, 0, 0)
	If StringRegExp($sFilter, "[\\/:><\|]|(?s)^\s*$") Then Return SetError(2, 0, 0)
	If Not ($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 0, 0)
	Local $hSearch = FileFindFirstFile($sFilePath & $sFilter)
	If @error Then Return SetError(4, 0, 0)
	While 1
		$sFileName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If ($iFlag + @extended = 2) Then ContinueLoop
		$sFileList &= $sDelimiter & $sFullPath & $sFileName
	WEnd
	FileClose($hSearch)
	If $sFileList = "" Then Return SetError(4, 0, 0)
	Return StringSplit(StringTrimLeft($sFileList, 1), $sDelimiter)
EndFunc   ;==>_FileListToArray
Func _FileListToArrayRec($sFilePath, $sMask = "*", $iReturn = $FLTAR_FILESFOLDERS, $iRecur = $FLTAR_NORECUR, $iSort = $FLTAR_NOSORT, $iReturnPath = $FLTAR_RELPATH)
	If Not FileExists($sFilePath) Then Return SetError(1, 1, "")
	If $sMask = Default Then $sMask = "*"
	If $iReturn = Default Then $iReturn = $FLTAR_FILESFOLDERS
	If $iRecur = Default Then $iRecur = $FLTAR_NORECUR
	If $iSort = Default Then $iSort = $FLTAR_NOSORT
	If $iReturnPath = Default Then $iReturnPath = $FLTAR_RELPATH
	If $iRecur > 1 Or Not IsInt($iRecur) Then Return SetError(1, 6, "")
	Local $bLongPath = False
	If StringLeft($sFilePath, 4) == "\\?\" Then
		$bLongPath = True
	EndIf
	Local $sFolderSlash = ""
	If StringRight($sFilePath, 1) = "\" Then
		$sFolderSlash = "\"
	Else
		$sFilePath = $sFilePath & "\"
	EndIf
	Local $asFolderSearchList[100] = [1]
	$asFolderSearchList[1] = $sFilePath
	Local $iHide_HS = 0, $sHide_HS = ""
	If BitAND($iReturn, 4) Then
		$iHide_HS += 2
		$sHide_HS &= "H"
		$iReturn -= 4
	EndIf
	If BitAND($iReturn, 8) Then
		$iHide_HS += 4
		$sHide_HS &= "S"
		$iReturn -= 8
	EndIf
	Local $iHide_Link = 0
	If BitAND($iReturn, 16) Then
		$iHide_Link = 0x400
		$iReturn -= 16
	EndIf
	Local $iMaxLevel = 0
	If $iRecur < 0 Then
		StringReplace($sFilePath, "\", "", 0, $STR_NOCASESENSEBASIC)
		$iMaxLevel = @extended - $iRecur
	EndIf
	Local $sExclude_List = "", $sExclude_List_Folder = "", $sInclude_List = "*"
	Local $aMaskSplit = StringSplit($sMask, "|")
	Switch $aMaskSplit[0]
		Case 3
			$sExclude_List_Folder = $aMaskSplit[3]
			ContinueCase
		Case 2
			$sExclude_List = $aMaskSplit[2]
			ContinueCase
		Case 1
			$sInclude_List = $aMaskSplit[1]
	EndSwitch
	Local $sInclude_File_Mask = ".+"
	If $sInclude_List <> "*" Then
		If Not __FLTAR_ListToMask($sInclude_File_Mask, $sInclude_List) Then Return SetError(1, 2, "")
	EndIf
	Local $sInclude_Folder_Mask = ".+"
	Switch $iReturn
		Case 0
			Switch $iRecur
				Case 0
					$sInclude_Folder_Mask = $sInclude_File_Mask
			EndSwitch
		Case 2
			$sInclude_Folder_Mask = $sInclude_File_Mask
	EndSwitch
	Local $sExclude_File_Mask = ":"
	If $sExclude_List <> "" Then
		If Not __FLTAR_ListToMask($sExclude_File_Mask, $sExclude_List) Then Return SetError(1, 3, "")
	EndIf
	Local $sExclude_Folder_Mask = ":"
	If $iRecur Then
		If $sExclude_List_Folder Then
			If Not __FLTAR_ListToMask($sExclude_Folder_Mask, $sExclude_List_Folder) Then Return SetError(1, 4, "")
		EndIf
		If $iReturn = 2 Then
			$sExclude_Folder_Mask = $sExclude_File_Mask
		EndIf
	Else
		$sExclude_Folder_Mask = $sExclude_File_Mask
	EndIf
	If Not ($iReturn = 0 Or $iReturn = 1 Or $iReturn = 2) Then Return SetError(1, 5, "")
	If Not ($iSort = 0 Or $iSort = 1 Or $iSort = 2) Then Return SetError(1, 7, "")
	If Not ($iReturnPath = 0 Or $iReturnPath = 1 Or $iReturnPath = 2) Then Return SetError(1, 8, "")
	If $iHide_Link Then
		Local $tFile_Data = DllStructCreate("struct;align 4;dword FileAttributes;uint64 CreationTime;uint64 LastAccessTime;uint64 LastWriteTime;" & "dword FileSizeHigh;dword FileSizeLow;dword Reserved0;dword Reserved1;wchar FileName[260];wchar AlternateFileName[14];endstruct")
		Local $hDLL = DllOpen('kernel32.dll'), $aDLL_Ret
	EndIf
	Local $asReturnList[100] = [0]
	Local $asFileMatchList = $asReturnList, $asRootFileMatchList = $asReturnList, $asFolderMatchList = $asReturnList
	Local $bFolder = False, $hSearch = 0, $sCurrentPath = "", $sName = "", $sRetPath = ""
	Local $iAttribs = 0, $sAttribs = ''
	Local $asFolderFileSectionList[100][2] = [[0, 0]]
	While $asFolderSearchList[0] > 0
		$sCurrentPath = $asFolderSearchList[$asFolderSearchList[0]]
		$asFolderSearchList[0] -= 1
		Switch $iReturnPath
			Case 1
				$sRetPath = StringReplace($sCurrentPath, $sFilePath, "")
			Case 2
				If $bLongPath Then
					$sRetPath = StringTrimLeft($sCurrentPath, 4)
				Else
					$sRetPath = $sCurrentPath
				EndIf
		EndSwitch
		If $iHide_Link Then
			$aDLL_Ret = DllCall($hDLL, 'handle', 'FindFirstFileW', 'wstr', $sCurrentPath & "*", 'struct*', $tFile_Data)
			If @error Or Not $aDLL_Ret[0] Then
				ContinueLoop
			EndIf
			$hSearch = $aDLL_Ret[0]
		Else
			$hSearch = FileFindFirstFile($sCurrentPath & "*")
			If $hSearch = -1 Then
				ContinueLoop
			EndIf
		EndIf
		If $iReturn = 0 And $iSort And $iReturnPath Then
			__FLTAR_AddToList($asFolderFileSectionList, $sRetPath, $asFileMatchList[0] + 1)
		EndIf
		$sAttribs = ''
		While 1
			If $iHide_Link Then
				$aDLL_Ret = DllCall($hDLL, 'int', 'FindNextFileW', 'handle', $hSearch, 'struct*', $tFile_Data)
				If @error Or Not $aDLL_Ret[0] Then
					ExitLoop
				EndIf
				$sName = DllStructGetData($tFile_Data, "FileName")
				If $sName = ".." Then
					ContinueLoop
				EndIf
				$iAttribs = DllStructGetData($tFile_Data, "FileAttributes")
				If $iHide_HS And BitAND($iAttribs, $iHide_HS) Then
					ContinueLoop
				EndIf
				If BitAND($iAttribs, $iHide_Link) Then
					ContinueLoop
				EndIf
				$bFolder = False
				If BitAND($iAttribs, 16) Then
					$bFolder = True
				EndIf
			Else
				$bFolder = False
				$sName = FileFindNextFile($hSearch, 1)
				If @error Then
					ExitLoop
				EndIf
				$sAttribs = @extended
				If StringInStr($sAttribs, "D") Then
					$bFolder = True
				EndIf
				If StringRegExp($sAttribs, "[" & $sHide_HS & "]") Then
					ContinueLoop
				EndIf
			EndIf
			If $bFolder Then
				Select
					Case $iRecur < 0
						StringReplace($sCurrentPath, "\", "", 0, $STR_NOCASESENSEBASIC)
						If @extended < $iMaxLevel Then
							ContinueCase
						EndIf
					Case $iRecur = 1
						If Not StringRegExp($sName, $sExclude_Folder_Mask) Then
							__FLTAR_AddToList($asFolderSearchList, $sCurrentPath & $sName & "\")
						EndIf
				EndSelect
			EndIf
			If $iSort Then
				If $bFolder Then
					If StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
						__FLTAR_AddToList($asFolderMatchList, $sRetPath & $sName & $sFolderSlash)
					EndIf
				Else
					If StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
						If $sCurrentPath = $sFilePath Then
							__FLTAR_AddToList($asRootFileMatchList, $sRetPath & $sName)
						Else
							__FLTAR_AddToList($asFileMatchList, $sRetPath & $sName)
						EndIf
					EndIf
				EndIf
			Else
				If $bFolder Then
					If $iReturn <> 1 And StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
						__FLTAR_AddToList($asReturnList, $sRetPath & $sName & $sFolderSlash)
					EndIf
				Else
					If $iReturn <> 2 And StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
						__FLTAR_AddToList($asReturnList, $sRetPath & $sName)
					EndIf
				EndIf
			EndIf
		WEnd
		If $iHide_Link Then
			DllCall($hDLL, 'int', 'FindClose', 'ptr', $hSearch)
		Else
			FileClose($hSearch)
		EndIf
	WEnd
	If $iHide_Link Then
		DllClose($hDLL)
	EndIf
	If $iSort Then
		Switch $iReturn
			Case 2
				If $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
				ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
				$asReturnList = $asFolderMatchList
				__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
			Case 1
				If $asRootFileMatchList[0] = 0 And $asFileMatchList[0] = 0 Then Return SetError(1, 9, "")
				If $iReturnPath = 0 Then
					__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
					__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
				Else
					__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList, 1)
				EndIf
			Case 0
				If $asRootFileMatchList[0] = 0 And $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
				If $iReturnPath = 0 Then
					__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
					$asReturnList[0] += $asFolderMatchList[0]
					ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
					_ArrayConcatenate($asReturnList, $asFolderMatchList, 1)
					__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
				Else
					Local $asReturnList[$asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0] + 1]
					$asReturnList[0] = $asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0]
					__ArrayDualPivotSort($asRootFileMatchList, 1, $asRootFileMatchList[0])
					For $i = 1 To $asRootFileMatchList[0]
						$asReturnList[$i] = $asRootFileMatchList[$i]
					Next
					Local $iNextInsertionIndex = $asRootFileMatchList[0] + 1
					__ArrayDualPivotSort($asFolderMatchList, 1, $asFolderMatchList[0])
					Local $sFolderToFind = ""
					For $i = 1 To $asFolderMatchList[0]
						$asReturnList[$iNextInsertionIndex] = $asFolderMatchList[$i]
						$iNextInsertionIndex += 1
						If $sFolderSlash Then
							$sFolderToFind = $asFolderMatchList[$i]
						Else
							$sFolderToFind = $asFolderMatchList[$i] & "\"
						EndIf
						Local $iFileSectionEndIndex = 0, $iFileSectionStartIndex = 0
						For $j = 1 To $asFolderFileSectionList[0][0]
							If $sFolderToFind = $asFolderFileSectionList[$j][0] Then
								$iFileSectionStartIndex = $asFolderFileSectionList[$j][1]
								If $j = $asFolderFileSectionList[0][0] Then
									$iFileSectionEndIndex = $asFileMatchList[0]
								Else
									$iFileSectionEndIndex = $asFolderFileSectionList[$j + 1][1] - 1
								EndIf
								If $iSort = 1 Then
									__ArrayDualPivotSort($asFileMatchList, $iFileSectionStartIndex, $iFileSectionEndIndex)
								EndIf
								For $k = $iFileSectionStartIndex To $iFileSectionEndIndex
									$asReturnList[$iNextInsertionIndex] = $asFileMatchList[$k]
									$iNextInsertionIndex += 1
								Next
								ExitLoop
							EndIf
						Next
					Next
				EndIf
		EndSwitch
	Else
		If $asReturnList[0] = 0 Then Return SetError(1, 9, "")
		ReDim $asReturnList[$asReturnList[0] + 1]
	EndIf
	Return $asReturnList
EndFunc   ;==>_FileListToArrayRec
Func __FLTAR_AddFileLists(ByRef $asTarget, $asSource_1, $asSource_2, $iSort = 0)
	ReDim $asSource_1[$asSource_1[0] + 1]
	If $iSort = 1 Then __ArrayDualPivotSort($asSource_1, 1, $asSource_1[0])
	$asTarget = $asSource_1
	$asTarget[0] += $asSource_2[0]
	ReDim $asSource_2[$asSource_2[0] + 1]
	If $iSort = 1 Then __ArrayDualPivotSort($asSource_2, 1, $asSource_2[0])
	_ArrayConcatenate($asTarget, $asSource_2, 1)
EndFunc   ;==>__FLTAR_AddFileLists
Func __FLTAR_AddToList(ByRef $aList, $vValue_0, $vValue_1 = -1)
	If $vValue_1 = -1 Then
		$aList[0] += 1
		If UBound($aList) <= $aList[0] Then ReDim $aList[UBound($aList) * 2]
		$aList[$aList[0]] = $vValue_0
	Else
		$aList[0][0] += 1
		If UBound($aList) <= $aList[0][0] Then ReDim $aList[UBound($aList) * 2][2]
		$aList[$aList[0][0]][0] = $vValue_0
		$aList[$aList[0][0]][1] = $vValue_1
	EndIf
EndFunc   ;==>__FLTAR_AddToList
Func __FLTAR_ListToMask(ByRef $sMask, $sList)
	If StringRegExp($sList, "\\|/|:|\<|\>|\|") Then Return 0
	$sList = StringReplace(StringStripWS(StringRegExpReplace($sList, "\s*;\s*", ";"), $STR_STRIPLEADING + $STR_STRIPTRAILING), ";", "|")
	$sList = StringReplace(StringReplace(StringRegExpReplace($sList, "[][$^.{}()+\-]", "\\$0"), "?", "."), "*", ".*?")
	$sMask = "(?i)^(" & $sList & ")\z"
	Return 1
EndFunc   ;==>__FLTAR_ListToMask
Func _PathFull($sRelativePath, $sBasePath = @WorkingDir)
	If Not $sRelativePath Or $sRelativePath = "." Then Return $sBasePath
	Local $sFullPath = StringReplace($sRelativePath, "/", "\")
	Local Const $sFullPathConst = $sFullPath
	Local $sPath
	Local $bRootOnly = StringLeft($sFullPath, 1) = "\" And StringMid($sFullPath, 2, 1) <> "\"
	If $sBasePath = Default Then $sBasePath = @WorkingDir
	For $i = 1 To 2
		$sPath = StringLeft($sFullPath, 2)
		If $sPath = "\\" Then
			$sFullPath = StringTrimLeft($sFullPath, 2)
			Local $nServerLen = StringInStr($sFullPath, "\") - 1
			$sPath = "\\" & StringLeft($sFullPath, $nServerLen)
			$sFullPath = StringTrimLeft($sFullPath, $nServerLen)
			ExitLoop
		ElseIf StringRight($sPath, 1) = ":" Then
			$sFullPath = StringTrimLeft($sFullPath, 2)
			ExitLoop
		Else
			$sFullPath = $sBasePath & "\" & $sFullPath
		EndIf
	Next
	If StringLeft($sFullPath, 1) <> "\" Then
		If StringLeft($sFullPathConst, 2) = StringLeft($sBasePath, 2) Then
			$sFullPath = $sBasePath & "\" & $sFullPath
		Else
			$sFullPath = "\" & $sFullPath
		EndIf
	EndIf
	Local $aTemp = StringSplit($sFullPath, "\")
	Local $aPathParts[$aTemp[0]], $j = 0
	For $i = 2 To $aTemp[0]
		If $aTemp[$i] = ".." Then
			If $j Then $j -= 1
		ElseIf Not ($aTemp[$i] = "" And $i <> $aTemp[0]) And $aTemp[$i] <> "." Then
			$aPathParts[$j] = $aTemp[$i]
			$j += 1
		EndIf
	Next
	$sFullPath = $sPath
	If Not $bRootOnly Then
		For $i = 0 To $j - 1
			$sFullPath &= "\" & $aPathParts[$i]
		Next
	Else
		$sFullPath &= $sFullPathConst
		If StringInStr($sFullPath, "..") Then $sFullPath = _PathFull($sFullPath)
	EndIf
	Do
		$sFullPath = StringReplace($sFullPath, ".\", "\")
	Until @extended = 0
	Return $sFullPath
EndFunc   ;==>_PathFull
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" & ((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func _WinAPI_GetString($pString, $bUnicode = True)
	Local $iLength = _WinAPI_StrLen($pString, $bUnicode)
	If @error Or Not $iLength Then Return SetError(@error + 10, @extended, '')
	Local $tString = DllStructCreate(__Iif($bUnicode, 'wchar', 'char') & '[' & ($iLength + 1) & ']', $pString)
	If @error Then Return SetError(@error, @extended, '')
	Return SetExtended($iLength, DllStructGetData($tString, 1))
EndFunc   ;==>_WinAPI_GetString
Func _WinAPI_PathIsDirectory($sFilePath)
	Local $aRet = DllCall('shlwapi.dll', 'bool', 'PathIsDirectoryW', 'wstr', $sFilePath)
	If @error Then Return SetError(@error, @extended, False)
	Return $aRet[0]
EndFunc   ;==>_WinAPI_PathIsDirectory
Func _WinAPI_StrLen($pString, $bUnicode = True)
	Local $W = ''
	If $bUnicode Then $W = 'W'
	Local $aRet = DllCall('kernel32.dll', 'int', 'lstrlen' & $W, 'struct*', $pString)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aRet[0]
EndFunc   ;==>_WinAPI_StrLen
Func __Inc(ByRef $aData, $iIncrement = 100)
	Select
		Case UBound($aData, $UBOUND_COLUMNS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
			Else
				$aData[0][0] += 1
				If $aData[0][0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
				EndIf
			EndIf
		Case UBound($aData, $UBOUND_ROWS)
			If $iIncrement < 0 Then
				ReDim $aData[$aData[0] + 1]
			Else
				$aData[0] += 1
				If $aData[0] > UBound($aData) - 1 Then
					ReDim $aData[$aData[0] + $iIncrement]
				EndIf
			EndIf
		Case Else
			Return 0
	EndSelect
	Return 1
EndFunc   ;==>__Inc
Func __Iif($bTest, $vTrue, $vFalse)
	Return $bTest ? $vTrue : $vFalse
EndFunc   ;==>__Iif
Func __WINVER()
	Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
	DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
	Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
	If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
	Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc   ;==>__WINVER
Func _WinAPI_CommandLineToArgv($sCmd)
	Local $aResult[1] = [0]
	$sCmd = StringStripWS($sCmd, $STR_STRIPLEADING + $STR_STRIPTRAILING)
	If Not $sCmd Then
		Return $aResult
	EndIf
	Local $aRet = DllCall('shell32.dll', 'ptr', 'CommandLineToArgvW', 'wstr', $sCmd, 'int*', 0)
	If @error Or Not $aRet[0] Or (Not $aRet[2]) Then Return SetError(@error + 10, @extended, 0)
	Local $tPtr = DllStructCreate('ptr[' & $aRet[2] & ']', $aRet[0])
	Dim $aResult[$aRet[2] + 1] = [$aRet[2]]
	For $i = 1 To $aRet[2]
		$aResult[$i] = _WinAPI_GetString(DllStructGetData($tPtr, 1, $i))
	Next
	DllCall("kernel32.dll", "handle", "LocalFree", "handle", $aRet[0])
	Return $aResult
EndFunc   ;==>_WinAPI_CommandLineToArgv
Func _WinAPI_PathAddBackslash($sFilePath)
	Local $tPath = DllStructCreate('wchar[260]')
	DllStructSetData($tPath, 1, $sFilePath)
	Local $aRet = DllCall('shlwapi.dll', 'ptr', 'PathAddBackslashW', 'struct*', $tPath)
	If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
	Return DllStructGetData($tPath, 1)
EndFunc   ;==>_WinAPI_PathAddBackslash
Global Const $tagPRINTDLG = __Iif(@AutoItX64, '', 'align 2;') & 'dword Size;hwnd hOwner;handle hDevMode;handle hDevNames;handle hDC;dword Flags;word FromPage;word ToPage;word MinPage;word MaxPage;word Copies;handle hInstance;lparam lParam;ptr PrintHook;ptr SetupHook;ptr PrintTemplateName;ptr SetupTemplateName;handle hPrintTemplate;handle hSetupTemplate'
Func _WinAPI_ChangeWindowMessageFilterEx($hWnd, $iMsg, $iAction)
	Local $tCFS, $aRet
	If $hWnd And ($__WINVER > 0x0600) Then
		Local Const $tagCHANGEFILTERSTRUCT = 'dword cbSize; dword ExtStatus'
		$tCFS = DllStructCreate($tagCHANGEFILTERSTRUCT)
		DllStructSetData($tCFS, 1, DllStructGetSize($tCFS))
		$aRet = DllCall('user32.dll', 'bool', 'ChangeWindowMessageFilterEx', 'hwnd', $hWnd, 'uint', $iMsg, 'dword', $iAction, 'struct*', $tCFS)
	Else
		$tCFS = 0
		$aRet = DllCall('user32.dll', 'bool', 'ChangeWindowMessageFilter', 'uint', $iMsg, 'dword', $iAction)
	EndIf
	If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
	Return SetExtended(DllStructGetData($tCFS, 2), 1)
EndFunc   ;==>_WinAPI_ChangeWindowMessageFilterEx
Func _WinAPI_DragQueryFileEx($hDrop, $iFlag = 0)
	Local $aRet = DllCall('shell32.dll', 'uint', 'DragQueryFileW', 'handle', $hDrop, 'uint', -1, 'ptr', 0, 'uint', 0)
	If @error Then Return SetError(@error, @extended, 0)
	If Not $aRet[0] Then Return SetError(10, 0, 0)
	Local $iCount = $aRet[0]
	Local $aResult[$iCount + 1]
	For $i = 0 To $iCount - 1
		$aRet = DllCall('shell32.dll', 'uint', 'DragQueryFileW', 'handle', $hDrop, 'uint', $i, 'wstr', '', 'uint', 4096)
		If Not $aRet[0] Then Return SetError(11, 0, 0)
		If $iFlag Then
			Local $bDir = _WinAPI_PathIsDirectory($aRet[3])
			If (($iFlag = 1) And $bDir) Or (($iFlag = 2) And Not $bDir) Then
				ContinueLoop
			EndIf
		EndIf
		$aResult[$i + 1] = $aRet[3]
		$aResult[0] += 1
	Next
	If Not $aResult[0] Then Return SetError(12, 0, 0)
	__Inc($aResult, -1)
	Return $aResult
EndFunc   ;==>_WinAPI_DragQueryFileEx
#EndRegion ###include###

Global Const $sAppRegKey = "HKEY_CURRENT_USER\Software\Trong\AttributesTool"
Global Const $sKeyREGexp = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Global Const $OSArch = @OSArch
Global Const $WindowsDir = @WindowsDir
Global Const $TempDir = @TempDir
Global Const $HomeDrive = @HomeDrive
Global Const $UserProfileDir = @UserProfileDir
Global Const $AppDataDir = @AppDataDir
Global Const $AppDataCommonDir = @AppDataCommonDir
Local $ComSpecX = @ComSpec
If StringInStr($OSArch, "64") Then
	DllCall("kernel32.dll", "boolean", "Wow64EnableWow64FsRedirection", "boolean", 0)
	If FileExists($WindowsDir & "\System32\cmd.exe") Then $ComSpecX = '"' & $WindowsDir & '\System32\cmd.exe"'
EndIf
;~ 	If StringInStr($OSArch, "64") Then DllCall("kernel32.dll", "boolean", "Wow64EnableWow64FsRedirection", "boolean", 1)
Global Const $ComSpec = $ComSpecX
Global Const $sOption1 = "-RAHS", $sOption2 = "+RAHS", $sOption3 = "+RH", $sOption4 = "+RS", $sOption5 = "+SH"
Global Const $sOption6 = "+R", $sOption7 = "+S", $sOption8 = "+H", $sOption9 = "Delete", $sOption10 = "Rename Unicode"
Global Const $sOption11 = "Set files/folders timestamp:", $sOption12 = "+C", $sOption13 = "+M", $sOption14 = "+A"
Global $AppTitle = "Attributes Tool"
#Region # GUI
Global $AppWindows = GUICreate($AppTitle, 564, 79, -1, -1, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlCreateLabel("Select Option > ", 4, 20, 91, 17, BitOR($SS_CENTER, $SS_CENTERIMAGE), $GUI_WS_EX_PARENTDRAG)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt1 = GUICtrlCreateRadio("-RAHS", 103, 9, 55, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Bỏ thuộc tính: Chỉ Đọc,Lưu Trữ,Ẩn và Hệ Thống", "Clear Attributes: ReadOnly,Archive,Hidden and System", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt2 = GUICtrlCreateRadio("+RAHS", 163, 9, 55, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Chỉ Đọc,Lưu Trữ,Ẩn và Hệ Thống", "Set Attributes: ReadOnly,Archive,Hidden and System", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt3 = GUICtrlCreateRadio("+RH", 223, 9, 42, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Chỉ Đọc và Ẩn", "Set Attributes: ReadOnly and Hidden", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt4 = GUICtrlCreateRadio("+RS", 268, 9, 42, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Chỉ Đọc và Hệ Thống", "Set Attributes: ReadOnly and System", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt5 = GUICtrlCreateRadio("+SH", 315, 9, 42, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Ẩn và Hệ Thống", "Set Attributes: Hidden and System", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt6 = GUICtrlCreateRadio("+R", 365, 9, 33, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Chỉ Đọc", "Set Attributes: ReadOnly", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt7 = GUICtrlCreateRadio("+S", 407, 9, 33, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Hệ Thống", "Set Attributes: System", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt8 = GUICtrlCreateRadio("+H", 447, 9, 33, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Đặt thuộc tính: Ẩn", "Set Attributes: Hidden", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt9 = GUICtrlCreateRadio("Delete", 102, 56, 55, 17)
GUICtrlSetTip(-1, "Xóa tập tin và thư mục!", "Delete Files/Folders", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt10 = GUICtrlCreateRadio("Rename Unicode", 160, 56, 105, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Loại bỏ ký tự Unicode (dấu Tiếng Việt) ở tên tập tin/thư mục!", "Removing Unicode characters in name of Files/Folders!", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $rOpt11 = GUICtrlCreateRadio("Set files/folders timestamp:", 102, 32, 148, 17)
GUICtrlSetTip(-1, "Đặt thời gian TẠO/SỬA/TRUY CẬP cho tập tin và thư mục", "Set time CREATED/MODIFIED/ACCESSED for files/folders", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $iTimestamp = GUICtrlCreateInput(@YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC, 250, 32, 110, 19)
GUICtrlSetTip(-1, 'The new time to set in the format "YYYYMMDDHHMMSS" ' & @CRLF & '(Year, month, day, hours (24hr clock), minutes, seconds). ' & @CRLF & 'If the time is blank then the current time is used.', 'Định dạng: NămThángNgàyGiờPhútGiây (Thêm dấu 0 trước Số nhỏ hơn 10)', 1, 1)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $cCreated = GUICtrlCreateCheckbox("+C", 366, 32, 33, 17)
GUICtrlSetTip(-1, "Đặt giờ Tạo!", "The timestamp to change: Created", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
Global $cModified = GUICtrlCreateCheckbox("+M", 408, 32, 33, 17)
GUICtrlSetTip(-1, "Đặt giờ lần thay đổi cuối cùng!", "The timestamp to change: Last modified", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
Global $cAccessed = GUICtrlCreateCheckbox("+A", 448, 32, 33, 17)
GUICtrlSetTip(-1, "Đặt giờ lần truy cập/mở cuối cùng!", "The timestamp to change: Last accessed", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
Global $cRecurseFolder = GUICtrlCreateCheckbox("Recurse Folder", 267, 56, 90, 17, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Áp dụng bao gồm cả các tập tin và thư mục con!", "Includes files/folder on subfolders.", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
Global $bFixAutorunUSB = GUICtrlCreateButton("USB Hidden Folder Fix", 362, 52, 115, 25, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Tự động xóa/sửa lỗi và hiện file bị ẩn do virus trên USB", "Fix hidden file/folder on USB", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $bSelectFileIN = GUICtrlCreateButton("Drag and Drop Files/Folders HERE", 482, 8, 79, 69, BitOR($BS_VCENTER, $BS_MULTILINE, $BS_FLAT))
GUICtrlSetTip(-1, "Kéo và thả các tập tin và thư mục muốn áp dụng vào đây!", "Drag and Drop Files/Folders HERE", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $bRemoveTemp = GUICtrlCreateButton("R", 16, 40, 15, 15, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Dọn dẹp file tạm trong windows!", "Remove template files", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $bAbout = GUICtrlCreateButton("A", 32, 40, 15, 15, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Thông tin tác giả!", "About", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $bHideFile = GUICtrlCreateButton("H", 48, 40, 15, 15, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Ẩn tập tin ẨN và thư mục ẨN trên Windows!", "Hide Hidden File/Folder", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $bShowFile = GUICtrlCreateButton("S", 64, 40, 15, 15, BitOR($BS_VCENTER, $BS_FLAT))
GUICtrlSetTip(-1, "Hiện tập tin ẨN và thư mục ẨN trên Windows!", "Show Hidden File/Folder", 1, 1)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
Global $cForce = GUICtrlCreateCheckbox("Try Force", 16, 60, 65, 17, BitOR($GUI_SS_DEFAULT_CHECKBOX, $BS_VCENTER))
GUICtrlSetTip(-1, "Cố gắng thực hiện lệnh!", "Try to enforce the the command!", 1, 1)
Global $pStatus = GUICtrlCreateProgress(0, 0, 565, 5, $PBS_SMOOTH)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
WinSetTrans($AppWindows, "", 236)
#EndRegion # GUI
Global $__aDropFiles, $sPercent
GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES")
GUISetState(@SW_SHOW, $AppWindows)
_WinAPI_ChangeWindowMessageFilterEx($AppWindows, $WM_DROPFILES, $MSGFLT_ALLOW)
_WinAPI_ChangeWindowMessageFilterEx($AppWindows, $WM_COPYDATA, $MSGFLT_ALLOW)
_WinAPI_ChangeWindowMessageFilterEx($AppWindows, $WM_COPYGLOBALDATA, $MSGFLT_ALLOW)
GUICtrlSetData($pStatus, 100)
Global $sOption = RegRead($sAppRegKey, "Mode")
Global $sRecurse = RegRead($sAppRegKey, "Recurse")
If @error Or $sRecurse <> 0 Then GUICtrlSetState($cRecurseFolder, $GUI_CHECKED)
Global $sTimestamp = RegRead($sAppRegKey, "Timestamp")
If (StringLen($sTimestamp) <> 14) Or (Not StringIsInt($sTimestamp)) Then $sTimestamp = ""
GUICtrlSetData($iTimestamp, $sTimestamp)
Global $sCreated = RegRead($sAppRegKey, "Created")
If @error Or $sCreated <> 0 Then GUICtrlSetState($cCreated, $GUI_CHECKED)
Global $sModified = RegRead($sAppRegKey, "Modified")
If @error Or $sModified <> 0 Then GUICtrlSetState($cModified, $GUI_CHECKED)
Global $sAccessed = RegRead($sAppRegKey, "Accessed")
If @error Or $sAccessed <> 0 Then GUICtrlSetState($cAccessed, $GUI_CHECKED)
Global $sForce = RegRead($sAppRegKey, "TryForce")
If $sForce = 1 Then GUICtrlSetState($cForce, $GUI_CHECKED)
If $sOption = $sOption11 Then
	GUICtrlSetState($rOpt11, $GUI_CHECKED)
ElseIf $sOption = $sOption10 Then
	GUICtrlSetState($rOpt10, $GUI_CHECKED)
ElseIf $sOption = $sOption9 Then
	GUICtrlSetState($rOpt9, $GUI_CHECKED)
ElseIf $sOption = $sOption8 Then
	GUICtrlSetState($rOpt8, $GUI_CHECKED)
ElseIf $sOption = $sOption7 Then
	GUICtrlSetState($rOpt7, $GUI_CHECKED)
ElseIf $sOption = $sOption6 Then
	GUICtrlSetState($rOpt6, $GUI_CHECKED)
ElseIf $sOption = $sOption5 Then
	GUICtrlSetState($rOpt5, $GUI_CHECKED)
ElseIf $sOption = $sOption4 Then
	GUICtrlSetState($rOpt4, $GUI_CHECKED)
ElseIf $sOption = $sOption3 Then
	GUICtrlSetState($rOpt3, $GUI_CHECKED)
ElseIf $sOption = $sOption2 Then
	GUICtrlSetState($rOpt2, $GUI_CHECKED)
Else
	$sOption = $sOption1
	GUICtrlSetState($rOpt1, $GUI_CHECKED)
EndIf
Global $nMsg, $sPercent, $iPercent, $sMode, $isONE
Local $aCmdLineRaw = StringReplace($CmdLineRaw, '/ErrorStdOut "' & @ScriptFullPath & '"', "")
Local $aCmdLine = _WinAPI_CommandLineToArgv($aCmdLineRaw)
If IsArray($aCmdLine) And $aCmdLine[0] > 0 Then
	_GetSetStatus()
	If $sMode = 3 Then
		If StringInStr($sTimestamp, "Invalid") Then
			MsgBox(48, "Time is Invalid: Thời gian sai định dạng!", 'The new time to set in the format "YYYYMMDDHHMMSS" ' & @CRLF & '(Year, month, day, hours (24hr clock), minutes, seconds). ' & @CRLF & 'If the time is blank then the current time is used.' & @CRLF & ' Using a date earlier than 1980-01-01 will have no effect.' & @CRLF & @CRLF & 'Định dạng: NămThángNgàyGiờPhútGiây ' & @CRLF & '(Thêm dấu 0 trước Số nhỏ hơn 10)' & @CRLF & 'Thời gian phải lớn hơn 1980-01-01', Default, $AppWindows)
			Exit
		EndIf
	EndIf
	WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
	If $aCmdLine[0] = 1 Then
		_DoIt($aCmdLine[1], 1, 1, 1)
	Else
		For $i = 1 To $aCmdLine[0]
			_DoIt($aCmdLine[$i], $i, $aCmdLine[0])
		Next
	EndIf
	GUICtrlSetData($pStatus, 100)
	WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
	Exit
Else
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $rOpt1, $rOpt2, $rOpt3, $rOpt4, $rOpt5, $rOpt6, $rOpt7, $rOpt8, $rOpt9, $rOpt10, $rOpt11, $cRecurseFolder, $iTimestamp, $cCreated, $cModified, $cAccessed, $cForce
				_GetSetStatus()
			Case $GUI_EVENT_DROPPED
				_GetSetStatus()
				If $sMode = 3 Then
					If StringInStr($sTimestamp, "Invalid") Then
						MsgBox(48, "Time is Invalid: Thời gian sai định dạng!", 'The new time to set in the format "YYYYMMDDHHMMSS" ' & @CRLF & '(Year, month, day, hours (24hr clock), minutes, seconds). ' & @CRLF & 'If the time is blank then the current time is used.' & @CRLF & ' Using a date earlier than 1980-01-01 will have no effect.' & @CRLF & @CRLF & 'Định dạng: NămThángNgàyGiờPhútGiây ' & @CRLF & '(Thêm dấu 0 trước Số nhỏ hơn 10)' & @CRLF & 'Thời gian phải lớn hơn 1980-01-01', Default, $AppWindows)
						ContinueLoop
					EndIf
				EndIf
				WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
				If $__aDropFiles[0] > 0 Then
					If $__aDropFiles[0] = 1 Then
						_DoIt($__aDropFiles[1], 1, 1, 1)
					Else
						For $i = 1 To $__aDropFiles[0]
							_DoIt($__aDropFiles[$i], $i, $__aDropFiles[0])
						Next
					EndIf
				EndIf
				GUICtrlSetData($pStatus, 100)
				WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
			Case $bSelectFileIN
				Local $zListFileIN, $zFileIN = FileOpenDialog("Select Files to Apply", @WorkingDir, "All Files(*.*)", $FD_MULTISELECT, "", $AppWindows)
				If Not @error Then
					_GetSetStatus()
					If $sMode = 3 Then
						If StringInStr($sTimestamp, "Invalid") Then
							MsgBox(48, "Time is Invalid: Thời gian sai định dạng!", 'The new time to set in the format "YYYYMMDDHHMMSS" ' & @CRLF & '(Year, month, day, hours (24hr clock), minutes, seconds). ' & @CRLF & 'If the time is blank then the current time is used. ' & @CRLF & 'Using a date earlier than 1980-01-01 will have no effect.' & @CRLF & @CRLF & 'Định dạng: NămThángNgàyGiờPhútGiây ' & @CRLF & '(Thêm dấu 0 trước Số nhỏ hơn 10)' & @CRLF & 'Thời gian phải lớn hơn 1980-01-01', Default, $AppWindows)
							ContinueLoop
						EndIf
					EndIf
					WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
					If StringInStr($zFileIN, "|") Then
						$zListFileIN = StringSplit($zFileIN, "|")
						If IsArray($zListFileIN) Then
							For $i = 2 To $zListFileIN[0]
								_DoIt($zListFileIN[1] & "\" & $zListFileIN[$i], $i, $zListFileIN[0])
							Next
						EndIf
					Else
						_DoIt($zFileIN, 1, 1, 1)
					EndIf
				EndIf
				GUICtrlSetData($pStatus, 100)
				WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
			Case $bAbout
				MsgBox(64, "Attributes Tool: About", "Attributes Tool: v2.0" & @CRLF & @CRLF & @YEAR & " © Ðào Văn Trong - Trong.CF", Default, $AppWindows)
			Case $bShowFile
				_ShowHiddenFile(1)
			Case $bHideFile
				_ShowHiddenFile(0)
			Case $bRemoveTemp
				WinSetTitle($AppWindows, "", $AppTitle & " - " & " Cleaning....")
				_RemoveTemp()
				GUICtrlSetData($pStatus, 100)
				WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
			Case $bFixAutorunUSB
				_FixUSB()
				GUICtrlSetData($pStatus, 100)
				WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
			Case $GUI_EVENT_CLOSE
				_GetSetStatus()
				WinSetTitle($AppWindows, "", $AppTitle & " - " & "Setting Saved!")
				For $i = 255 To 0 Step -10
					WinSetTrans($AppWindows, "", $i)
				Next
				GUIDelete($AppWindows)
				Exit
		EndSwitch
	WEnd
EndIf

#Region # Function
Func _DoIt($sFileIN, $iNum, $tNum, $inONE = 0)
	$sPercent = Int(($iNum / $tNum) * 10000) / 100
	If $sPercent > 100 Then $sPercent = 100
	If $inONE = 1 Then $sPercent = $sPercent / 2
	GUICtrlSetData($pStatus, $sPercent)
	WinSetTitle($AppWindows, "", $AppTitle & " - " & $sFileIN)
	If $sMode = 1 Then
		_DeleteIT($sFileIN, $sRecurse)
	ElseIf $sMode = 2 Then
		_RenameUnicode($sFileIN, $sRecurse)
	ElseIf $sMode = 3 Then
		If StringInStr($sTimestamp, "Invalid") Then MsgBox(48, $AppTitle, "Time is Invalid")
		_SetTimestamp($sFileIN, $sRecurse)
	Else
		_SetAttributes($sFileIN, $sOption, $sRecurse)
	EndIf
	Sleep(100)
EndFunc   ;==>_DoIt
Func _RemoveTemp()
	MsgBox(48, "Warring!", "Please save your work and close all web session before Press OK", Default, $AppWindows)
	Local $xBefore = DriveSpaceFree($HomeDrive)
	Local $xTimer = TimerInit()
	FileRecycleEmpty()
	Local $vNum = 100
	While $vNum
		ProcessClose("opera.exe")
		If ProcessExists("opera.exe") = 0 Then ExitLoop
		$vNum -= 1
	WEnd
	$vNum = 100
	While $vNum
		ProcessClose("firefox.exe")
		If ProcessExists("firefox.exe") = 0 Then ExitLoop
		$vNum -= 1
	WEnd
	$vNum = 100
	While $vNum
		ProcessClose("chrome.exe")
		If ProcessExists("chrome.exe") = 0 Then ExitLoop
		$vNum -= 1
	WEnd
	$vNum = 100
	While $vNum
		ProcessClose("old_chrome.exe")
		If ProcessExists("old_chrome.exe") = 0 Then ExitLoop
		$vNum -= 1
	WEnd
	_DeleteIT($TempDir, 1, 1)
	_DeleteIT($WindowsDir & "\Temp", 1, 1)
	_DeleteIT($WindowsDir & "\System32\config\systemprofile\Local Settings\Temp", 1, 1)
	_DeleteIT($WindowsDir & "\System32\config\systemprofile\Local Settings\Temporary Internet Files", 1, 1)
	_DeleteIT($WindowsDir & "\SoftwareDistribution\Download", 1, 1)
	_DeleteIT($WindowsDir & "\ie8updates", 1, 0)
	_DeleteIT($HomeDrive & "\$WINDOWS.~BT", 1, 0)
	_DeleteIT($HomeDrive & "\Windows.old", 1, 0)
	_DeleteIT($HomeDrive & "\swsetup", 1, 0)
	_DeleteIT($WindowsDir & "\$hf_mig$", 1, 0)
	_DeleteIT($HomeDrive & "\Temp", 1, 1)
	_DeleteIT($HomeDrive & "\Temp\Temporary Internet Files", 1, 1)
	_DeleteIT($UserProfileDir & "\Local Settings\Temp", 1, 1)
	_DeleteIT($UserProfileDir & "\Local Settings\History", 1, 1)
	_DeleteIT($UserProfileDir & "\Local Settings\Temporary Internet Files", 1, 1)
	_DeleteIT($UserProfileDir & "\Recent", 1, 1)
	_DeleteIT($UserProfileDir & "\Cookies", 1, 1)
	_DeleteIT($AppDataDir & "\Microsoft\Office\Recent", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\Local\Temp", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\LocalLow\Temp", 1, 1)
	_DeleteIT($UserProfileDir & "\Appdata\Local\Microsoft\Windows\History", 1, 1)
	_DeleteIT($UserProfileDir & "\Appdata\Local\Microsoft\Windows\WER\ReportArchive", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\Local\Microsoft\Windows\Temporary Internet Files", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\Local\Microsoft\Windows\INetCache", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\Local\Microsoft\Windows\INetCookies", 1, 1)
	_DeleteIT($UserProfileDir & "\Appdata\Local\Microsoft\Terminal Server Client\Cache", 1, 1)
	_DeleteIT($AppDataCommonDir & "\Microsoft\Windows\WER\ReportArchive", 1, 1)
	_DeleteIT($AppDataCommonDir & "\Microsoft\Windows Defender\Scans\History\Results\Resource", 1, 1)
	_DeleteIT($AppDataDir & "\Sun\Java\Deployment\Cache", 1, 1)
	_DeleteIT($AppDataDir & "\Microsoft\Windows\Cookies", 1, 1)
	_DeleteIT($AppDataDir & "\Microsoft\Windows\Recent", 1, 1)
	_DeleteIT($AppDataDir & "\Microsoft\Office\Recent", 1, 1)
	_DeleteIT($UserProfileDir & "\AppData\Local\Opera Software\Opera Stable\Cache", 1, 1)
	_DeleteIT($UserProfileDir & "\Local Settings\Application Data\Opera Software\Opera Stable\Cache", 1, 1)
	Local $sListFile = _FileListToArray($UserProfileDir & "\Local Settings\Application Data\Mozilla\Firefox\Profiles", "*", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		If FileExists($sListFile[$x] & "\Cache\*.*") Then
			_DeleteIT($sListFile[$x] & "\Cache", 1, 1)
			_DeleteIT($sListFile[$x] & "\Cache2", 1, 1)
			_DeleteIT($sListFile[$x] & "\thumbnails", 1, 1)
			_DeleteIT($sListFile[$x] & "\startupCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\jumpListCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\OfflineCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\cookies.sqlite", 1, 1)
		EndIf
	Next
	Local $sListFile = _FileListToArray($UserProfileDir & "\Appdata\Local\Mozilla\Firefox\Profiles", "*", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		If FileExists($sListFile[$x] & "\Cache\*.*") Then
			_DeleteIT($sListFile[$x] & "\Cache", 1, 1)
			_DeleteIT($sListFile[$x] & "\Cache2", 1, 1)
			_DeleteIT($sListFile[$x] & "\thumbnails", 1, 1)
			_DeleteIT($sListFile[$x] & "\startupCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\jumpListCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\OfflineCache", 1, 1)
			_DeleteIT($sListFile[$x] & "\cookies.sqlite", 1, 1)
		EndIf
	Next
	Local $sListFile = _FileListToArray($UserProfileDir & "\AppData\Local\Google\Chrome\User Data", "*", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		If FileExists($sListFile[$x] & "\Cache\*.*") Then
			_DeleteIT($sListFile[$x] & "\Cache", 1, 1)
			_DeleteIT($sListFile[$x] & "\Cookies-journal", 1, 1)
			_DeleteIT($sListFile[$x] & "\Current Session", 1, 1)
			_DeleteIT($sListFile[$x] & "\History", 1, 1)
			_DeleteIT($sListFile[$x] & "\History-journal", 1, 1)
		EndIf
	Next
	Local $sListFile = _FileListToArray($UserProfileDir & "\Local Settings\Application Data\Google\Chrome\User Data", "*", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		If FileExists($sListFile[$x] & "\Cache\*.*") Then
			_DeleteIT($sListFile[$x] & "\Cache", 1, 1)
			_DeleteIT($sListFile[$x] & "\Cookies-journal", 1, 1)
			_DeleteIT($sListFile[$x] & "\Current Session", 1, 1)
			_DeleteIT($sListFile[$x] & "\History", 1, 1)
			_DeleteIT($sListFile[$x] & "\History-journal", 1, 1)
		EndIf
	Next
	_DeleteIT(@ProgramFilesDir & "\Google\Update\Download", 1, 1)
	_DeleteIT(@ProgramFilesDir & "\Google\Update\Install", 1, 1)
	_DeleteIT(@ProgramFilesDir & "\Google\Chrome\Application\old_chrome.exe", 1, 0)
	Local $xListFile = _FileListToArray(@ProgramFilesDir & "\Google\Chrome\Application", "*", 2, 0)
	Local $xLargestVersion
	For $x = 0 To UBound($xListFile) - 1
		StringReplace($xListFile[$x], ".", "")
		If @extended > 2 Then $xLargestVersion &= $xListFile[$x] & "|"
	Next
	If StringRight($xLargestVersion, 1) = "|" Then $xLargestVersion = StringTrimRight($xLargestVersion, 1)
	Local $xArray = StringSplit($xLargestVersion, "|")
	_ArraySort($xArray, 0, 0, 0, 0, "_CompareVersion")
	Local $sLargestVersion = $xArray[UBound($xArray) - 1]
	For $x = 1 To UBound($xArray) - 1
		If $xArray[$x] <> $sLargestVersion Then _DeleteIT(@ProgramFilesDir & "\Google\Chrome\Application\" & $xArray[$x], 1, 0)
	Next
	Local $sListFile = _FileListToArray($WindowsDir, "$*Uninstall*$", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		_DeleteIT($sListFile[$x], 1, 0)
	Next
	Local $sListFile = _FileListToArray($WindowsDir, "KB*.log", 2, 1)
	For $x = 1 To UBound($sListFile) - 1
		_DeleteIT($sListFile[$x], 1, 0)
	Next
	WinSetTitle($AppWindows, "", $AppTitle & " - " & " Cleaning....")
	RunWait($ComSpec & ' /c Dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase', "", @SW_MINIMIZE)
	RunWait($ComSpec & ' /c Dism.exe /online /Cleanup-Image /SPSuperseded', "", @SW_MINIMIZE)
	ShellExecuteWait("RunDll32.exe", " InetCpl.cpl,ClearMyTracksByProcess 255")
	_ClipEmpty()
	Local $xAfter = DriveSpaceFree($HomeDrive)
	Local $xFinishTimer = TimerDiff($xTimer)
	WinSetTitle($AppWindows, "", $AppTitle & " - " & " All Done!")
	MsgBox(64, "Finished", Round($xAfter - $xBefore, 1) & "MB has been removed" & @CRLF & "in " & Round(($xFinishTimer / 1000), 3) & " seconds!", Default, $AppWindows)
EndFunc   ;==>_RemoveTemp
Func _ClipEmpty()
	Local $Success = 0
	DllCall('user32.dll', 'int', 'OpenClipboard', 'hwnd', 0)
	If @error = 0 Then
		DllCall('user32.dll', 'int', 'EmptyClipboard')
		If @error = 0 Then $Success = 1
		DllCall('user32.dll', 'int', 'CloseClipboard')
		If $Success Then Return 1
	EndIf
	DllCall('user32.dll', 'int', 'CloseClipboard')
	Return 0
EndFunc   ;==>_ClipEmpty
Func _FixUSB()
	Local $aArray = DriveGetDrive($DT_ALL)
	If @error Then Return SetError(1, 0, 0)
	If $sForce = 1 Then _FixVirusUserInit()
	For $i = 1 To $aArray[0]
		Local $sDrive = StringUpper($aArray[$i])
		If DriveStatus($sDrive & "\") = $DS_READY Then
			If _IsUSB($sDrive) Then
				WinSetTitle($AppWindows, "", $AppTitle & " - Fix USB: " & "Please waitting...")
				Local $sListFile = _FileListToArray($sDrive, "*.*", 1, 1)
				For $x = 1 To UBound($sListFile) - 1
					FileSetAttrib($sListFile[$x], "-RASH", 0)
					If StringLower(_SplitPath($sListFile[$x], 4)) = ".lnk" Then _DeleteIT($sListFile[$x])
				Next
			EndIf
		EndIf
	Next
	For $i = 1 To $aArray[0]
		Local $sDrive = StringUpper($aArray[$i])
		If DriveStatus($sDrive & "\") = $DS_READY Then
			If _IsUSB($sDrive) Then
				WinSetTitle($AppWindows, "", $AppTitle & " - Fix USB: " & "Please waitting...")
				Local $sListFile = _FileListToArrayRec($sDrive, "*", 0, $sRecurse, 2, 2)
				For $x = 1 To UBound($sListFile) - 1
					Local $sFileDir = $sListFile[$x]
					FileSetAttrib($sFileDir, "-RASH")
					If FileExists(_SplitPath($sFileDir, 7)) Then
						If StringLower(_SplitPath($sFileDir, 4)) = ".exe" And _IsFile($sFileDir) Then _DeleteIT($sFileDir)
					EndIf
				Next
				If FileExists($sDrive & "\System Volume Information") Then _DeleteIT($sDrive & "\System Volume Information", 1)
				If FileExists($sDrive & "\RECYCLER") Then _DeleteIT($sDrive & "\RECYCLER", 1)
				FileWrite($sDrive & "\RECYCLER", "Protection from viruses!")
				FileSetAttrib($sDrive & "\RECYCLER", "+RASHO")
				If FileExists($sDrive & "\RECYCLER") Then _DeleteIT($sDrive & "\$RECYCLE.BIN", 1)
				FileWrite($sDrive & "\$RECYCLE.BIN", "Protection from viruses!")
				FileSetAttrib($sDrive & "\$RECYCLE.BIN", "+RASHO")
			EndIf
			Local $sDriverType = DriveGetType($sDrive)
			If $sDriverType = "Removable" Or $sDriverType = "Fixed" Then
				WinSetTitle($AppWindows, "", $AppTitle & " - Fix USB: " & "Please waitting...")
				Local $sListFile = _FileListToArray($sDrive, "FOUND.*", 0, 1)
				For $x = 1 To UBound($sListFile) - 1
					_DeleteIT($sListFile[$x], 1)
				Next
				Local $sProcessList = StringSplit("Images.exe|SysAnti.exe|SafeSys.exe|svcshot.exe|System.exe|phimnguoilon.exe|phim nguoi lon.exe|phimhot.exe|phim hot.exe|special.exe|secret.exe|bimat.exe|bi mat.exe|forever.exe|romantic.exe|mixa.exe|nhacmoi.exe|new folder.exe|newfolder.exe|image.exe|images.exe|My Music.exe|MyMusic.exe|Music.exe|my picture.exe|mypicture.exe|picture.exe", "|")
				For $v = 1 To UBound($sProcessList) - 1
					If _IsFile($sDrive & "\" & $sProcessList[$v]) Then
						_DeleteIT($sDrive & "\" & $sProcessList[$v])
					Else
						FileDelete($sDrive & "\" & $sProcessList[$v] & "\*.*")
						If FileExists($sDrive & "\" & $sProcessList[$v]) Then FileSetAttrib($sDrive & "\" & $sProcessList[$v], "+RASHO")
					EndIf
				Next
				If FileExists($sDrive & "\Autorun.bak") Then _DeleteIT($sDrive & "\Autorun.bak", 1)
				If FileExists($sDrive & "\Autorun.inf") Then
					If StringInStr(FileRead($sDrive & "\Autorun.inf"), "SETUP.exe") Then
						FileSetAttrib($sDrive & "\Autorun.inf", "-RASH", 0)
						FileMove($sDrive & "\Autorun.inf", $sDrive & "\Autorun.bak", 1)
						If @error And $sForce = 1 Then
							_TakeOwnership($sDrive & "\Autorun.inf")
							FileMove($sDrive & "\Autorun.inf", $sDrive & "\Autorun.bak", 1)
						EndIf
					EndIf
					If _IsFile($sDrive & "\Autorun.inf") Then _DeleteIT($sDrive & "\Autorun.inf", 1)
				EndIf
				DirCreate($sDrive & "\Autorun.inf")
				FileSetAttrib($sDrive & "\Autorun.inf", "+RASHO")
			EndIf
		EndIf
	Next
	_UpdateExplorer()
EndFunc   ;==>_FixUSB
Func _FixVirusUserInit()
	Local $sWinlogonKEY = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
	Local $dUserInit = RegRead($sWinlogonKEY, "Userinit")
	Local $dShell = RegRead($sWinlogonKEY, "Shell")
	Local $dUIHost = RegRead($sWinlogonKEY, "UIHost")
	Local $dShellInfrastructure = RegRead($sWinlogonKEY, "ShellInfrastructure")
	Local $sModifl = $dUIHost <> "" And StringLower($dUIHost) <> StringLower("LogonUI.exe")
	Local $sModifi = StringLower($dShell) <> StringLower("explorer.exe")
	Local $sModify = StringLower($dUserInit) <> StringLower("C:\WINDOWS\system32\userinit.exe,")
	Local $sModif1 = $dShellInfrastructure <> "" And StringLower($dShellInfrastructure) <> StringLower("sihost.exe")
	Local $sProcess, $sListProcess = StringSplit("Images.exe|SysAnti.exe|SafeSys.exe|svcshot.exe|System.exe|phimnguoilon.exe|phim nguoi lon.exe|phimhot.exe|phim hot.exe|special.exe|secret.exe|bimat.exe|bi mat.exe|forever.exe|romantic.exe|mixa.exe|nhacmoi.exe|new folder.exe|newfolder.exe|image.exe|images.exe|My Music.exe|MyMusic.exe|Music.exe|my picture.exe|mypicture.exe|picture.exe", "|")
	For $i = 1 To UBound($sListProcess) - 1
		If ProcessExists($sListProcess[$i]) Then $sProcess &= ' /IM "' & $sListProcess[$i] & '"'
	Next
	If $sProcess <> "" Then RunWait(@ComSpec & " /c taskkill /F " & $sProcess, "", @SW_HIDE)
	If $sModifl Or $sModifi Or $sModify Or $sModif1 Then
		RegWrite($sWinlogonKEY, "Userinit", "REG_SZ", "C:\WINDOWS\system32\userinit.exe,")
		RegWrite($sWinlogonKEY, "Shell", "REG_SZ", "explorer.exe")
		If $dUIHost <> "" Then RegWrite($sWinlogonKEY, "UIHost", "REG_EXPAND_SZ", "LogonUI.exe")
		If $dShellInfrastructure <> "" Then RegWrite($sWinlogonKEY, "ShellInfrastructure", "REG_SZ", "sihost.exe")
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies")
		_DeleteIT($WindowsDir & "\System32\system.exe")
		_DeleteIT($WindowsDir & "\System32\task.exe")
		_DeleteIT($WindowsDir & "\userinit.exe")
		_DeleteIT($WindowsDir & "\kdcoms.dll")
		_DeleteIT($WindowsDir & "\kdcoms32.dll")
	EndIf
EndFunc   ;==>_FixVirusUserInit
Func _IsUSB($sDrive)
	Local $sIsUSB = 0, $sType = DriveGetType($sDrive, 3)
	If DriveGetType($sDrive, 1) = "Removable" Or $sType = "USB" Or $sType = "SD" Or $sType = "MMC" Then $sIsUSB = 1
	Return $sIsUSB
EndFunc   ;==>_IsUSB
Func _ShowHiddenFile($iShow)
	If $iShow Then
		RegWrite($sKeyREGexp, "HideFileExt", "REG_DWORD", 0)
		RegWrite($sKeyREGexp, "Hidden", "REG_DWORD", 1)
		RegWrite($sKeyREGexp, "SuperHidden", "REG_DWORD", 1)
		RegWrite($sKeyREGexp, "ShowSuperHidden", "REG_DWORD", 1)
		MsgBox(0, $AppTitle, "Files/folders Hidden are displayed!" & @CRLF & _UpdateExplorer() & @CRLF & "Các tập tin/thư mục ẨN đã được hiển thị.", Default, $AppWindows)
	Else
		RegWrite($sKeyREGexp, "HideFileExt", "REG_DWORD", 0)
		RegWrite($sKeyREGexp, "Hidden", "REG_DWORD", 2)
		RegWrite($sKeyREGexp, "SuperHidden", "REG_DWORD", 0)
		RegWrite($sKeyREGexp, "ShowSuperHidden", "REG_DWORD", 0)
		MsgBox(0, $AppTitle, "Files/folders Hidden are hide!" & @CRLF & _UpdateExplorer() & @CRLF & "Các tập tin/thư mục ẨN đã được ẨN.", Default, $AppWindows)
	EndIf
EndFunc   ;==>_ShowHiddenFile
Func _UpdateExplorer()
	Local $bold = Opt("WinSearchChildren", True)
	Local $a = WinList("[CLASS:SHELLDLL_DefView]")
	For $i = 0 To UBound($a) - 1
		DllCall("user32.dll", "long", "SendMessage", "hwnd", $a[$i][1], "int", 273, "int", 28931, "int", 0)
	Next
	Opt("WinSearchChildren", $bold)
	Return ">"
EndFunc   ;==>_UpdateExplorer
Func _SetTimestamp($sFileDirIN, $sRecurseDir = 0)
	WinSetTitle($AppWindows, "", $AppTitle & " - Sets timestamp: " & $sFileDirIN)
	ConsoleRead("!" & $sTimestamp & @CRLF)
	If $sRecurseDir = 1 Then
		If StringInStr(FileGetAttrib($sFileDirIN), 'D') <> 1 Then $sRecurseDir = 0
	EndIf
	If $sModified = 1 Then FileSetTime($sFileDirIN, $sTimestamp, $FT_MODIFIED, $sRecurseDir)
	If $sCreated = 1 Then FileSetTime($sFileDirIN, $sTimestamp, $FT_CREATED, $sRecurseDir)
	If $sAccessed = 1 Then FileSetTime($sFileDirIN, $sTimestamp, $FT_ACCESSED, $sRecurseDir)
	Return 1
EndFunc   ;==>_SetTimestamp
Func _DeleteIT($sFileDirIN, $sRecurseDir = 0, $oEmtyDir = 0)
	If Not $oEmtyDir Then
		If $sRecurseDir = 1 Then
			If StringInStr(FileGetAttrib($sFileDirIN), 'D') <> 1 Then $sRecurseDir = 0
		EndIf
		If StringInStr($sFileDirIN, "*") Then
			Local $sTMPdir = _SplitPath($sFileDirIN, 6)
			Local $sTMPext = _SplitPath($sFileDirIN, 5)
			$sTMPext = StringReplace($sTMPext, "**", "*")
			If $sTMPext == "*.*" Then $sTMPext = "*"
			If $sRecurseDir Then
				WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
				Local $xList = _FileListToArray($sFileDirIN, $sTMPext, 0, 1)
				For $f = 0 To UBound($xList) - 1
					_DeleteIT($xList[$f], $sRecurseDir)
				Next
			EndIf
		Else
			If Not FileExists($sFileDirIN) Then Return SetError(1, 0, 0)
			WinSetTitle($AppWindows, "", $AppTitle & " - " & "Delete " & $sFileDirIN & " ...")
			If _IsFile($sFileDirIN) Then
				FileDelete($sFileDirIN)
				If FileExists($sFileDirIN) Then
					FileSetAttrib($sFileDirIN, "-RASH")
					If $sForce = 1 Then _TakeOwnership($sFileDirIN)
					If $sRecurseDir Then
						RunWait($ComSpec & ' /c DEL /F /S /Q "' & $sFileDirIN & '"', '', @SW_HIDE)
					Else
						RunWait($ComSpec & ' /c DEL /F /Q "' & $sFileDirIN & '"', '', @SW_HIDE)
					EndIf
					FileDelete($sFileDirIN)
				EndIf
				If FileExists($sFileDirIN) Then WinSetTitle($AppWindows, "", $AppTitle & " - " & "Error on delete " & $sFileDirIN)
			Else
				If StringInStr($sFileDirIN, ":") And StringLen($sFileDirIN) > 3 Then
					DirRemove($sFileDirIN, $sRecurseDir)
					If FileExists($sFileDirIN) Then
						FileSetAttrib($sFileDirIN, "-RASH", $sRecurseDir)
						If $sForce = 1 Then _TakeOwnership($sFileDirIN)
						DirRemove($sFileDirIN, $sRecurseDir)
						If FileExists($sFileDirIN) Then
							If $sRecurseDir Then
								RunWait($ComSpec & ' /c RD /S /Q "' & $sFileDirIN & '"', '', @SW_HIDE)
							Else
								RunWait($ComSpec & ' /c RD /Q "' & $sFileDirIN & '"', '', @SW_HIDE)
							EndIf
						EndIf
					EndIf
					If FileExists($sFileDirIN) Then WinSetTitle($AppWindows, "", $AppTitle & " - " & "Error on delete " & $sFileDirIN)
				Else
					If $sRecurseDir Then
						WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
						Local $xList = _FileListToArray($sFileDirIN, "*", 0, 1)
						For $c = 0 To UBound($xList) - 1
							_DeleteIT($xList[$c], $sRecurseDir)
						Next
						If FileExists(_WinAPI_PathAddBackslash($sFileDirIN) & '*.*') Then WinSetTitle($AppWindows, "", $AppTitle & " - " & "Error on delete " & $sFileDirIN)
					Else
						WinSetTitle($AppWindows, "", $AppTitle & " - " & "Error on delete " & $sFileDirIN)
					EndIf
				EndIf
			EndIf
		EndIf
	Else
		If $sRecurseDir And (Not _IsFile($sFileDirIN)) Then
			WinSetTitle($AppWindows, "", $AppTitle & " - " & "Please waitting...")
			Local $xList = _FileListToArray($sFileDirIN, "*", 0, 1)
			For $x = 0 To UBound($xList) - 1
				_DeleteIT($xList[$x], $sRecurseDir)
			Next
		EndIf
	EndIf
EndFunc   ;==>_DeleteIT
Func _IsFile($sPath)
	If Not FileExists($sPath) Then Return SetError(1, 0, -1)
	If StringInStr(FileGetAttrib($sPath), 'D') <> 0 Then
		Return SetError(0, 0, 0)
	Else
		Return SetError(0, 0, 1)
	EndIf
EndFunc   ;==>_IsFile
Func _RenameUnicode($sFileDirIN, $sRecurseDir)
	If Not FileExists($sFileDirIN) Then Return SetError(1, 0, 3)
	Local $sFileDirOUT
	If StringInStr(FileGetAttrib($sFileDirIN), 'D') <> 0 Then
		If $sRecurseDir Then
			Local $sFileList = _FileListToArrayRec($sFileDirIN, "*", 1, 1, 1, 2)
			For $x = 1 To UBound($sFileList) - 1
				GUICtrlSetData($pStatus, Int(($x / $sFileList[0]) * 10000) / 200)
				$sFileDirOUT = _SplitPath($sFileList[$x], 6) & _StringToASCII(_SplitPath($sFileList[$x], 5))
				_MoveFile($sFileList[$x], $sFileDirOUT)
			Next
			Local $sDirList = _FileListToArrayRec($sFileDirIN, "*", 2, 1, 1, 2)
			Local $nn, $nTotal = UBound($sDirList)
			For $x = 1 To $nTotal - 1
				GUICtrlSetData($pStatus, 50 + Int(($x / $sFileList[0]) * 10000) / 200)
				$nn = $nTotal - $x
				$sFileDirOUT = _SplitPath($sDirList[$nn], 6) & _StringToASCII(_SplitPath($sDirList[$nn], 3))
				_MoveDir($sDirList[$nn], $sFileDirOUT)
			Next
		EndIf
		$sFileDirOUT = _SplitPath($sFileDirIN, 6) & _StringToASCII(_SplitPath($sFileDirIN, 3))
		Return SetError(0, 0, _MoveDir($sFileDirIN, $sFileDirOUT))
	Else
		$sFileDirOUT = _SplitPath($sFileDirIN, 6) & _StringToASCII(_SplitPath($sFileDirIN, 5))
		Return SetError(0, 0, _MoveFile($sFileDirIN, $sFileDirOUT))
	EndIf
EndFunc   ;==>_RenameUnicode
Func _MoveDir($sFileDirIN, $sFileDirOUT)
	If Not FileExists($sFileDirOUT) Then
		DirMove($sFileDirIN, $sFileDirOUT, 1)
		If @error And $sForce = 1 Then
			_TakeOwnership($sFileDirIN)
			DirMove($sFileDirIN, $sFileDirOUT, 1)
		EndIf
		Return SetError(0, 0, FileExists($sFileDirOUT))
	EndIf
EndFunc   ;==>_MoveDir
Func _MoveFile($sFileDirIN, $sFileDirOUT)
	If Not FileExists($sFileDirOUT) Then
		FileMove($sFileDirIN, $sFileDirOUT, 1)
		If @error And $sForce = 1 Then
			_TakeOwnership($sFileDirIN)
			FileMove($sFileDirIN, $sFileDirOUT, 1)
		EndIf
		Return SetError(0, 0, FileExists($sFileDirOUT))
	EndIf
	Return SetError(0, 1, 0)
EndFunc   ;==>_MoveFile
Func _StringToASCII($xString)
	If $xString = "" Or StringIsASCII($xString) Then Return $xString
	Local $sCharVN[14][2] = [["á|à|ả|ã|ạ|â|ấ|ầ|ẩ|ẫ|ậ|ă|ắ|ằ|ẳ|ẵ|ặ", "a"], ["Á|À|Ả|Ã|Ạ|Â|Ấ|Ầ|Ẩ|Ẫ|Ậ|Ă|Ắ|Ằ|Ẳ|Ẵ|Ặ", "A"], ["đ", "d"], ["Đ", "D"], ["ê|é|è|ẻ|ẽ|ẹ|ế|ề|ể|ễ|ệ", "e"], ["Ê|É|È|Ẻ|Ẽ|Ẹ|Ế|Ề|Ể|Ễ|Ệ", "E"], ["ú|ù|ủ|ũ|ụ|ư|ứ|ừ|ử|ữ|ự", "u"], ["Ú|Ù|Ủ|Ũ|Ụ|Ư|Ứ|Ừ|Ử|Ữ|Ự", "U"], ["ó|ò|ỏ|õ|ọ|ô|ố|ồ|ổ|ỗ|ộ|ơ|ớ|ờ|ở|ỡ|ợ", "o"], ["Ó|Ò|Ỏ|Õ|Ọ|Ô|Ố|Ồ|Ổ|Ỗ|Ộ|Ơ|Ớ|Ờ|Ở|Ỡ|Ợ", "O"], ["í|ì|ỉ|ĩ|ị", "i"], ["Í|Ì|Ỉ|Ĩ|Ị", "I"], ["ý|ỳ|ỷ|ỹ|ỵ", "y"], ["Ý|Ỳ|Ỷ|Ỹ|Ỵ", "Y"]]
	For $i = 0 To 13
		$xString = StringRegExpReplace($xString, $sCharVN[$i][0], $sCharVN[$i][1])
	Next
	$xString = StringRegExpReplace($xString, '[[:^print:]]', "_")
	Do
		$xString = StringReplace($xString, "_-_", "_")
	Until Not StringInStr($xString, "_-_")
	Do
		$xString = StringReplace($xString, "__", "_")
	Until Not StringInStr($xString, "__")
	Return SetError(0, 0, $xString)
EndFunc   ;==>_StringToASCII
Func _SetAttributes($sFileDirIN, $sAttribSet = Default, $sRecurseDir = 1)
	If ($sAttribSet <> $sOption1 And $sAttribSet <> $sOption2 And $sAttribSet <> $sOption3 And $sAttribSet <> $sOption4 And $sAttribSet <> $sOption5 And $sAttribSet <> $sOption6 And $sAttribSet <> $sOption7 And $sAttribSet <> $sOption8) Or $sAttribSet = "" Or $sAttribSet = -1 Or $sAttribSet = Default Then $sAttribSet = $sOption1
	If ($sFileDirIN <> "") And ((Not StringInStr($sFileDirIN, "\")) Or (Not StringInStr($sFileDirIN, ":")) Or StringInStr($sFileDirIN, "..\")) Then $sFileDirIN = _PathFull($sFileDirIN, @WorkingDir)
	Local $SetAttrib, $sFileDirNameIN = _SplitPath($sFileDirIN, 5)
	If FileExists($sFileDirIN) Then
		If $sRecurseDir = 1 Then
			If StringInStr(FileGetAttrib($sFileDirIN), 'D') <> 1 Then $sRecurseDir = 0
		EndIf
		$SetAttrib = FileSetAttrib($sFileDirIN, $sAttribSet, $sRecurseDir)
		If (@error Or $SetAttrib = 0) And $sForce = 1 Then
			_TakeOwnership($sFileDirIN)
			$SetAttrib = FileSetAttrib($sFileDirIN, $sAttribSet, $sRecurseDir)
		EndIf
		If $SetAttrib <> 0 Then
			WinSetTitle($AppWindows, "", $AppTitle & " - " & "Set attribute " & $sAttribSet & " file " & $sFileDirNameIN & " success!")
		Else
			WinSetTitle($AppWindows, "", $AppTitle & " - " & "Set attribute " & $sAttribSet & " file " & $sFileDirNameIN & " failure!")
		EndIf
		Return SetError($SetAttrib = 0, 0, $SetAttrib > 0)
	Else
		WinSetTitle($AppWindows, "", $AppTitle & " - " & "File Not Found " & $sFileDirNameIN)
		Return SetError(1, 0, 0)
	EndIf
EndFunc   ;==>_SetAttributes
Func _TakeOwnership($xFile)
	If Not FileExists($xFile) Then Return SetError(1, 0, $xFile)
	If StringInStr(FileGetAttrib($xFile), 'D') <> 0 Then
		RunWait($ComSpec & ' /c takeown /f "' & $xFile & '" /R /D Y', '', @SW_HIDE)
		RunWait($ComSpec & ' /c Echo y|cacls "' & $xFile & '" /T /C /G Administrators:F', '', @SW_HIDE)
		RunWait($ComSpec & ' /c icacls "' & $xFile & '" /grant Administrators:F /T /C /Q', '', @SW_HIDE)
		Return SetError(0, 0, 0)
	Else
		RunWait($ComSpec & ' /c takeown /f "' & $xFile & '"', '', @SW_HIDE)
		RunWait($ComSpec & ' /c Echo y|cacls "' & $xFile & '" /C /G Administrators:F', '', @SW_HIDE)
		RunWait($ComSpec & ' /c icacls "' & $xFile & '" /grant Administrators:F /Q', '', @SW_HIDE)
		Return SetError(0, 0, 1)
	EndIf
	Return $xFile
EndFunc   ;==>_TakeOwnership
Func _GetSetStatus()
	$sOption = $sOption1
	$sMode = 0
	$sRecurse = 0
	$sTimestamp = ""
	$sCreated = -1
	$sModified = -1
	$sAccessed = -1
	$sForce = 0
	If GUICtrlRead($rOpt11) = $GUI_CHECKED Then
		$sOption = $sOption11
		$sMode = 3
	EndIf
	If GUICtrlRead($rOpt10) = $GUI_CHECKED Then
		$sOption = $sOption10
		$sMode = 2
	EndIf
	If GUICtrlRead($rOpt9) = $GUI_CHECKED Then
		$sOption = $sOption9
		$sMode = 1
	EndIf
	If GUICtrlRead($rOpt8) = $GUI_CHECKED Then $sOption = $sOption8
	If GUICtrlRead($rOpt7) = $GUI_CHECKED Then $sOption = $sOption7
	If GUICtrlRead($rOpt6) = $GUI_CHECKED Then $sOption = $sOption6
	If GUICtrlRead($rOpt5) = $GUI_CHECKED Then $sOption = $sOption5
	If GUICtrlRead($rOpt4) = $GUI_CHECKED Then $sOption = $sOption4
	If GUICtrlRead($rOpt3) = $GUI_CHECKED Then $sOption = $sOption3
	If GUICtrlRead($rOpt2) = $GUI_CHECKED Then $sOption = $sOption2
	If GUICtrlRead($cRecurseFolder) = $GUI_CHECKED Then $sRecurse = 1
	If GUICtrlRead($cForce) = $GUI_CHECKED Then $sForce = 1
	$sTimestamp = GUICtrlRead($iTimestamp, 0)
	If GUICtrlRead($cCreated) = $GUI_CHECKED Then $sCreated = 1
	If GUICtrlRead($cModified) = $GUI_CHECKED Then $sModified = 1
	If GUICtrlRead($cAccessed) = $GUI_CHECKED Then $sAccessed = 1
	RegWrite($sAppRegKey, "Created", "REG_DWORD", $sCreated)
	RegWrite($sAppRegKey, "Modified", "REG_DWORD", $sModified)
	RegWrite($sAppRegKey, "Accessed", "REG_DWORD", $sAccessed)
	RegWrite($sAppRegKey, "Timestamp", "REG_SZ", $sTimestamp)
	RegWrite($sAppRegKey, "Mode", "REG_SZ", $sOption)
	RegWrite($sAppRegKey, "Recurse", "REG_DWORD", $sRecurse)
	RegWrite($sAppRegKey, "TryForce", "REG_DWORD", $sForce)
	If $sTimestamp <> "" Then
		If (StringLen($sTimestamp) <> 14) Or (Not StringIsInt($sTimestamp)) Then $sTimestamp = "Time is Invalid"
	EndIf
	Return 1
EndFunc   ;==>_GetSetStatus
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
Func _SplitPath($sFilePath, $sType = 0)
	Local $sDrive, $sDir, $sFileName, $sExtension, $sReturn
	Local $aArray = StringRegExp($sFilePath, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", 1)
	If @error Then
		ReDim $aArray[5]
		$aArray[0] = $sFilePath
	EndIf
	$sDrive = $aArray[1]
	If StringLeft($aArray[2], 1) == "/" Then
		$sDir = StringRegExpReplace($aArray[2], "\h*[\/\\]+\h*", "\/")
	Else
		$sDir = StringRegExpReplace($aArray[2], "\h*[\/\\]+\h*", "\\")
	EndIf
	$aArray[2] = $sDir
	$sFileName = $aArray[3]
	$sExtension = $aArray[4]
	If $sType = 1 Then Return $sDrive
	If $sType = 2 Then Return $sDir
	If $sType = 3 Then Return $sFileName
	If $sType = 4 Then Return $sExtension
	If $sType = 5 Then Return $sFileName & $sExtension
	If $sType = 6 Then Return $sDrive & $sDir
	If $sType = 7 Then Return $sDrive & $sDir & $sFileName
	Return $aArray
EndFunc   ;==>_SplitPath
#EndRegion # Function

#Region # Author C490C3A06F2056C4836E2054726F6E67
