Attribute VB_Name = "Work02_Rotation2"


Function W_FindClaimPath2(productCode)
'Finance Case
    myStart = "\\vwaipth-share01\"
    Dim ws, ws02 As Worksheet
    Set ws = ThisWorkbook.Sheets("Master")
    Dim FNInfo, colB, colD, colE, TextFindRg, SearchRange As range
    Set ws02 = ThisWorkbook.Sheets("FolderName")
    Set colB = ws.Columns("B")
    Set colE = ws.Columns("E")
    Set colD = ws.Columns("D")

    If Left(productCode, 1) <> "6" Then
        FNStartStr = myStart & "Underwrite-Motor\LMG41\UnderWrite_Motor\08-Mgr.( SM Job Handover )\2) FN Campaign\SM\"
        On Error GoTo Err01
        
        Set SearchRange = ws02.usedRange
        Set TextFindRg = SearchRange.Find("!Finance", LookIn:=xlValues)
        Set TextFindRg = TextFindRg.Offset(1, 0)
        Set FNInfo = Sp_SelectFromTL(TextFindRg, 410, 11)
        'A_FindFromHookH1 ("* ï¿½ï¿½ SM3## ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½")
        myYear = VB_Xlookup(productCode, colB, colE)
        smName = VB_Xlookup(productCode, colB, colD)
        FullName = D_Xlookup2D(FNInfo, smName, myYear)
        
        out_str = FNStartStr & myYear & "\" & FullName
        W_FindClaimPath2 = out_str
        On Error GoTo 0
'Non Finance Case
    Else
        NFStartStr = myStart & "Underwrite-Motor\LMG41\UnderWrite_Motor\08-Mgr.( SM Job Handover )\3) NF Campaign"
        ' Change this later to make it more dynamic
        'If the range changes it could cause serious problems
        prodCodeList = ws02.range("W67:W679")
        FullNameList = ws02.range("X67:X679")
        FullName = VB_Xlookup(productCode, prodCodeList, FullNameList, "No Info")
        yr_Eng = VB_Xlookup(productCode, colB, colE, "No Info") - 543
        out_str = NFStartStr & "\" & yr_Eng & " NF campaign" & "\" & FullName
        W_FindClaimPath2 = out_str
            ' Change this later to make it more dynamic
        'If the range changes it could cause serious problems
    
    End If
    Exit Function
Err01:
    W_FindClaimPath2 = "Has no valid name for " & productCode
    'MsgBox ("Has no valid name for " & productCode)
End Function
Function W_FindClaimPath(productCode)
'Finance Case
    OtherStart = "Z:\LMG41\"
    Dim ws, ws02 As Worksheet
    Set ws = ThisWorkbook.Sheets("Master")
    Dim FNInfo, colB, colD, colE As range
    Set ws02 = ThisWorkbook.Sheets("FolderName")
    Set colB = ws.Columns("B")
    Set colE = ws.Columns("E")
    Set colD = ws.Columns("D")

    If Left(productCode, 1) <> "6" Then
        FNStartStr = OtherStart & "Underwrite-Motor\LMG41\UnderWrite_Motor\08-Mgr.( SM Job Handover )\2) FN Campaign\SM\"
        
        ' Change this later to make it more dynamic
        'If the range changes it could cause serious problems
        Set FNInfo = ws02.range("E186:N588")
        'A_FindFromHookH1 ("* ï¿½ï¿½ SM3## ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½")
        myYear = VB_Xlookup(productCode, colB, colE)
        smName = VB_Xlookup(productCode, colB, colD)
        FullName = D_Xlookup2D(FNInfo, smName, myYear)
        
        out_str = FNStartStr & myYear & "\" & FullName
        W_FindClaimPath = out_str
'Non Finance Case
    Else
        NFStartStr = OtherStart & "Underwrite-Motor\LMG41\UnderWrite_Motor\08-Mgr.( SM Job Handover )\3) NF Campaign"
        ' Change this later to make it more dynamic
        'If the range changes it could cause serious problems
        prodCodeList = ws02.range("W67:W679")
        FullNameList = ws02.range("X67:X679")
        FullName = VB_Xlookup(productCode, prodCodeList, FullNameList, "No Info")
        yr_Eng = VB_Xlookup(productCode, colB, colE, "No Info") - 543
        out_str = NFStartStr & "\" & yr_Eng & " NF campaign" & "\" & FullName
        W_FindClaimPath = out_str
            ' Change this later to make it more dynamic
        'If the range changes it could cause serious problems
    
    End If



    
End Function
Sub W_BlockLongText()

    'Declare variables for the worksheet and range
    Dim ws As Worksheet
    Dim rng As range
    
    'Set the worksheet variable to the active worksheet
    Set ws = ActiveSheet
    
    'Set the range variable to the selected range
    Set rng = selection
    
    'Loop through each cell in the range
    For Each cell In rng
        'Check the length of the cell's value
        If IsError(cell.value) Then
            cell.HorizontalAlignment = xlGeneral
        ElseIf Len(cell.value) > 10 Then
            'If the value is more than 10 characters, set the alignment to "Fill"
            cell.HorizontalAlignment = xlFill
        Else
            'If the value is less than or equal to 10 characters, set the alignment to "General"
            cell.HorizontalAlignment = xlGeneral
        End If
    Next cell

End Sub

Sub W_RearrangeByPrefix()
'worked
    Dim dataRange As range
    Dim prefixCol As Integer
    Set dataRange = selection
    prefixCol = 1 ' Assume prefix column is the leftmost column in the range
    Dim i As Integer
    Dim j As Integer
    Dim currentPrefix As String
    Dim currentValue As String
    Dim nextCol As Integer
    For i = 1 To dataRange.Rows.count
        currentPrefix = dataRange.Cells(i, prefixCol).value
        For j = 1 To dataRange.Columns.count
            If j <> prefixCol Then
                currentValue = dataRange.Cells(i, j).value
                If Left(currentValue, Len(currentPrefix)) <> currentPrefix And currentValue <> "" Then
                    ' Move value to new cell
                    nextRow = i + 1
                    dataRange.Cells(nextRow, j).value = currentValue
                    dataRange.Cells(i, j).value = ""
                End If
            End If
        Next j
    Next i
End Sub



Sub GetFolderPathSpace()
    path01 = "\\vwaipth-share01\Underwrite-Motor\LMG41\UnderWrite_Motor\08-Mgr.( SM Job Handover )\3) NF Campaign\2016 NF Campaign"
    Call OS_GetFolderNames(path01)
End Sub
Sub OS_GetExcelFileNames(folderPath)
    Dim fso As Object
    Dim folder As Object
    Dim File As Object
    Dim fileName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Dim out_arr() As Variant
    For Each File In folder.Files
        If LCase(Right(File.name, 4)) = ".xls" Or LCase(Right(File.name, 5)) = ".xlsx" Then
            fileName = Left(File.name, (Len(File.name) - 5))
            out_arr = A_Append(out_arr, fileName)
        End If
    Next File
    A_printArr (out_arr)
End Sub

Sub OS_GetFolderNames(path)
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    Dim out_arr() As Variant
    
    For Each subfolder In folder.SubFolders
        If subfolder.Attributes And vbDirectory Then
            folder_name = subfolder.name
            out_arr = A_Append(out_arr, folder_name)
        End If
    Next subfolder
    A_printArr (out_arr)

End Sub

Function Rg_Union(ParamArray rgs() As Variant) As range
'https://stackoverflow.com/questions/27554867/how-to-set-and-use-empty-range-in-vba
  Dim i As Long
  For i = 0 To UBound(rgs())
    If Not rgs(i) Is Nothing Then
      If Rg_Union Is Nothing Then Set Rg_Union = rgs(i) Else Set Rg_Union = Application.Union(Rg_Union, rgs(i))
    End If
  Next i
End Function
Function St_RemoveFileName(path)
    St_RemoveFileName = Left(path, InStrRev(path, "\", , vbTextCompare) - 1) & "\"
End Function

Function W_GetCarAge(text As String) As Variant
    Dim numbers() As String
    numbers = Split(text, "-")
    'à¸›à¹‰à¸²à¸¢à¹à¸”à¸‡ =>ï¿½ï¿½ï¿½ï¿½á´§
    If InStr(text, "»éÒÂá´§") > 0 Then
    'If InStr(Text, "à¸›à¹‰à¸²à¸¢à¹à¸”à¸‡") > 0 Then
        outArr = Array(1, 1)
        W_GetCarAge = outArr
        Exit Function
    End If
    
    If UBound(numbers) = 0 Then
        num = D_GetNum(text)
        W_GetCarAge = Array(num, num)
        Exit Function
    End If
    If UBound(numbers) = 1 Then
        num0 = D_GetNum(numbers(0))
        num1 = D_GetNum(numbers(1))
        W_GetCarAge = Array(num0, num1)
    Else
        W_GetCarAge = Array(0, 0)
    End If
End Function


Function D_GetNum(string_in) As Double
    Dim tempString As String
    tempString = ""
    For i = 1 To Len(string_in)
        If IsNumeric(Mid(string_in, i, 1)) Or Mid(string_in, i, 1) = "." Then
            tempString = tempString & Mid(string_in, i, 1)
        End If
    Next i
    If IsNumeric(tempString) Then
        D_GetNum = CDbl(tempString)
    Else
        D_GetNum = 0 ' return 0 if the string does not contain any numeric value
    End If
End Function

Sub A_FillValue(myArray As Variant, start_cell As Variant, Optional direction As XlDirection = xlDown, Optional direction2 = "", Optional ws_name = "", Optional overwrite = False, Optional time_option = False)
'More general and more powerful than A_printArr
'Done debugging for all cases(Could check more but it seems pretty stable from testing)
'I update case xlDown to make it faster
'But I haven't done for other filling * do that
'Hard for chatGPT
'time_option = False => No print time
'time_option = True => print time for debugging
    Dim rng As range
    Dim mySheet As Worksheet
    If ws_name = "" Then
        ws_name = ActiveSheet.name
    End If
    'Application.ScreenUpdating = False
    'Application.EnableEvents = False
    startTime = Timer
    Set mySheet = Worksheets(ws_name)
    
    If TypeName(start_cell) = "String" Then
        Set rng = mySheet.range(start_cell)
    Else
        Set rng = mySheet.range(start_cell.Address)
    End If
    
    If Not A_IsArray(myArray) Then
        rng.value = myArray
        Exit Sub
    End If
    If A_NDim(myArray) = 1 Then
    '1d array case
        Dim data() As Variant
            
        If direction = xlDown Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
            Set toFill = start_cell.Resize(UBound(myArray) - LBound(myArray) + 1, 1)
            ReDim data(LBound(myArray) To UBound(myArray), 1 To 1)
            For i = LBound(myArray) To UBound(myArray)
                data(i, 1) = myArray(i)
            Next i
            'For i = LBound(myArray) To UBound(myArray)
                'start_cell.Offset(i, 0).value = myArray(i)
            'Next i
            toFill.value = data
        ElseIf direction = xlUp Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
            For i = LBound(myArray) To UBound(myArray)
                start_cell.Offset(-i, 0).value = myArray(i)
            Next i
            
            'For i = LBound(myArray) To UBound(myArray)
                'new_i = UBound(myArray) - i
                'data(new_i, 1) = myArray(i)
            'Next i
        ElseIf direction = xlToRight Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
            For i = LBound(myArray) To UBound(myArray)
                start_cell.Offset(0, i).value = myArray(i)
            Next i
        ElseIf direction = xlToLeft Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
            For i = LBound(myArray) To UBound(myArray)
                start_cell.Offset(0, -i).value = myArray(i)
            Next i
        Else
            MsgBox "Invalid direction"
        End If
'For 2d Array case
    ElseIf A_NDim(myArray) = 2 Then
        If direction = xlDown Then
            If direction2 = "" Or direction2 = xlToRight Then
            'assume that index start with 0 otherwise it won't work
            'Because Offset needs 0 in order for it to select that start_cell
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(j, i).value = myArray(i, j)
                    Next j
                Next i
            ElseIf direction2 = xlToLeft Then
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToLeft))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(j, -i).value = myArray(i, j)
                    Next j
                Next i
            Else
                MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
            End If
        ElseIf direction = xlUp Then
            If direction2 = "" Or direction2 = xlToRight Then
            'assume that index start with 0 otherwise it won't work
            'Because Offset needs 0 in order for it to select that start_cell
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(-j, i).value = myArray(i, j)
                    Next j
                Next i
            ElseIf direction2 = xlToLeft Then
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToLeft))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(-j, -i).value = myArray(i, j)
                    Next j
                Next i
            Else
                MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
            End If
        ElseIf direction = xlToRight Then
            If direction2 = "" Or direction2 = xlDown Then
            'assume that index start with 0 otherwise it won't work
            'Because Offset needs 0 in order for it to select that start_cell
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(i, j).value = myArray(i, j)
                    Next j
                Next i
            ElseIf direction2 = xlUp Then
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlUp))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(-i, j).value = myArray(i, j)
                    Next j
                Next i
            Else
                MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
            End If
        ElseIf direction = xlToLeft Then
            If direction2 = "" Or direction2 = xlDown Then
            'assume that index start with 0 otherwise it won't work
            'Because Offset needs 0 in order for it to select that start_cell
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(i, -j).value = myArray(i, j)
                    Next j
                Next i
            ElseIf direction2 = xlUp Then
                Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlUp))
                For i = 0 To UBound(myArray, 1)
                    For j = 0 To UBound(myArray, 2)
                        start_cell.Offset(-i, -j).value = myArray(i, j)
                    Next j
                Next i
            Else
                MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
            End If
        Else
            MsgBox "Invalid direction"
        End If
    
    Else
        MsgBox ("Not support array with dimesion >2")
    End If
    'Application.ScreenUpdating = True
    'Application.EnableEvents = True
    endTime = Timer
    exetime = Format(endTime - startTime, "#0.00")
    If time_option Then
        MsgBox ("Time taken: " & exetime & " seconds")
    End If
    
End Sub
'thisWBName = "04.01  VBA Play V04"

Function A_FindFromHook(search_list, Optional offset_row = 1, Optional offset_col = 0, Optional ws_name)
'search_list could be string or Array
' If there are many cells that have the same word it will get only the 1st one
    Dim outArr As Variant
    If IsArray(search_list) Then
        For i = LBound(search_list) To UBound(search_list)
            curr_str = search_list(i)
            foundArr = A_FindFromHookH1(curr_str, offset_row, offset_col, ws_name)
            outArr = A_Extend(outArr, foundArr)
        Next
    Else
        outArr = A_FindFromHookH1(search_list, offset_row, offset_col, ws_name)
    End If
    A_FindFromHook = outArr
End Function

Function A_FindFromHookH1(search_str, Optional offset_row = 1, Optional offset_col = 0, Optional ws_name = "")
    Dim ws As Worksheet
    Dim ans_arr() As Variant
    If ws_name = "" Then
        ws_name = ActiveSheet.name
    End If
    Set ws = Worksheets(ws_name)

    Dim SearchRange As range
    Set SearchRange = ws.usedRange
    
    Dim text_cell As range
    Set text_cell = SearchRange.Find(What:=search_str, LookIn:=xlValues)
    'when it found nothing
    If text_cell Is Nothing Then
        A_FindFromHookH1 = ans_arr
        Exit Function
    End If
    
    Set target_start = text_cell.Offset(offset_row, offset_col)
    If target_start.value = "" Then
    'This could cause error
        A_FindFromHookH1 = "Nothing Found"
        Exit Function
    End If
    'Set row_used = F_UsedCellRow_Intersect(14)

    Set ans_cell = range(target_start, target_start.End(xlDown))
    ans_arr = A_toArray1d(ans_cell)
    A_FindFromHookH1 = ans_arr
    
End Function

Function W_FindUsedCellRight(cell) As range
    Dim ws As Worksheet
    Set ws = cell.Worksheet

    Dim SearchRange As range
    Set SearchRange = ws.range(cell.Offset(0, 1), ws.Cells(cell.row, ws.Columns.count).End(xlToLeft))

    Dim foundRange As range
    Set foundRange = SearchRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)

    If Not foundRange Is Nothing Then
        Set W_FindUsedCellRight = foundRange
    Else
        Set W_FindUsedCellRight = Nothing
    End If
End Function



Function F_UsedCellRow(n As Long) As range
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or set ws = ThisWorkbook.Sheets("Sheet1")

    Dim SearchRange As range
    Set SearchRange = ws.range("A" & n & ":XFD" & n) ' search from A to XFD (last column)

    Dim foundRange As range
    Set foundRange = SearchRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)

    If Not foundRange Is Nothing Then
        Set F_UsedCellRow = foundRange
    Else
        Set F_UsedCellRow = Nothing
    End If
End Function

Function F_UsedCellRow_Intersect(n As Long) As range
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or set ws = ThisWorkbook.Sheets("Sheet1")

    Dim rowRange As range
    Set rowRange = ws.range("A" & n & ":XFD" & n) ' range covering entire row

    Dim usedRange As range
    Set usedRange = ws.usedRange

    Dim intersectRange As range
    Set intersectRange = Intersect(rowRange, usedRange)

    If Not intersectRange Is Nothing Then
        Set F_UsedCellRow_Intersect = intersectRange.Cells(1, 1)
    Else
        Set F_UsedCellRow_Intersect = Nothing
    End If
End Function
Function Pl_Make(ws_name As String)
'THAI FIX
    wordArr = Array("Make", "ÂÕèËéÍ")
    outArr = A_FindFromHook(wordArr, , , ws_name)
    Pl_Make = outArr
    
'Pl = Pull data
End Function

Function Pl_Model(ws_name As String)
    wordArr = Array("Model", "ÃØè¹")
    Pl_Model = A_FindFromHook(wordArr, , , ws_name)
End Function
Function Pl_SumAssuredMin(ws_name As String)
' à¸—à¸¸à¸™à¸›à¸£à¸°à¸à¸±à¸™ => ï¿½Ø¹ï¿½ï¿½Ð¡Ñ¹
    wordArr = Array("·Ø¹»ÃÐ¡Ñ¹")
    Pl_SumAssuredMin = A_FindFromHook(wordArr, 1, 0, ws_name)
End Function
Function Pl_SumAssuredMax(ws_name As String)
    wordArr = Array("·Ø¹»ÃÐ¡Ñ¹")
    Pl_SumAssuredMax = A_FindFromHook(wordArr, 1, 2, ws_name)
End Function

Function Pl_MotorCode(ws_name As String)
'à¸£à¸«à¸±à¸ª => ï¿½ï¿½ï¿½ï¿½
    Dim ws As Worksheet
    Dim SearchRange, text_cell As range
    Set ws = Worksheets(ws_name)
    'search_str = "à¸£à¸«à¸±à¸ª"
    search_str = "ÃËÑÊ"
    'Set SearchRange = ws.usedRange
    Set text_cell = Rg_FindAllRangeH1(search_str, ws_name)
    Dim outArr() As Variant
    Dim myText As String
    For Each cell In text_cell
        myText = cell.value
        motorCode = D_GetNum(myText)
        If Not A_isInArr(outArr, motorCode) Then
            outArr = A_Append2(outArr, motorCode)
        End If
        
    Next
    Pl_MotorCode = outArr
End Function
Function Pl_Garage(ws_name As String)
'à¸‹à¹ˆà¸­à¸¡ => ï¿½ï¿½ï¿½ï¿½
    Dim garageRangeAll, garageRangeRes As range
    Set garageRangeAll = Rg_FindAllRange("«èÍÁ", ws_name)
'Assume there are 2 sections of motor code
    n_motorCode = garageRangeAll.Areas.count \ 2
    For i = 1 To n_motorCode
        Set garageRangeRes = Rg_Union(garageRangeAll.Areas(i), garageRangeRes)
    Next
    temp_arr = A_toArray1d(garageRangeRes)
    Dim convert_res() As Variant
    For Each x In temp_arr
    'ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½
        If x = "«èÍÁÍÙè" Then
        'If x = "à¸‹à¹ˆà¸­à¸¡à¸­à¸¹à¹ˆ" Then
            convert_res = A_Append2(convert_res, "Insurer")
    'ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
        'ElseIf x = "à¸‹à¹ˆà¸­à¸¡à¸«à¹‰à¸²à¸‡" Then
        ElseIf x = "«èÍÁËéÒ§" Then
            convert_res = A_Append2(convert_res, "Dealer")
        End If
    Next
    Pl_Garage = convert_res
    

End Function
Function W_GetInsurerCell1(ws_name As String)
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the top with age&motor code
    Dim str_arr() As Variant
    'ï¿½ï¿½ï¿½ï¿½á´§, ï¿½ï¿½ï¿½ï¿½, ï¿½ï¿½
    'str_arr = Array("Age", "Year")
    Set garageRangeAll = Rg_FindAllRange("«èÍÁÍÙè", ws_name)
    'Set garageRangeAll = Rg_FindAllRange("à¸‹à¹ˆà¸­à¸¡à¸­à¸¹à¹ˆ", ws_name)
    n_cell = garageRangeAll.Areas.count
'Assume that it has only 2 sets of ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
'Assume that it's non-contiguous
'Be careful with the case of contiguous
    n_motor = n_cell \ 2
    If n_cell = 1 Then
        Set W_GetInsurerCell1 = garageRangeAll
        Exit Function
    End If
    Dim outRng As range
    For i = 1 To n_motor
        Set outRng = Rg_Union(garageRangeAll.Areas(i), outRng)
    Next
    Set W_GetInsurerCell1 = outRng
End Function
Function W_GetInsurerCell2(ws_name As String)
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the top with age&motor code
    Dim str_arr() As Variant
    'ï¿½ï¿½ï¿½ï¿½á´§, ï¿½ï¿½ï¿½ï¿½, ï¿½ï¿½
    'str_arr = Array("Age", "Year")
    Set garageRangeAll = Rg_FindAllRange("«èÍÁÍÙè", ws_name)
    'Set garageRangeAll = Rg_FindAllRange("à¸‹à¹ˆà¸­à¸¡à¸­à¸¹à¹ˆ", ws_name)
    n_cell = garageRangeAll.Areas.count
'Assume that it has only 2 sets of ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
'Assume that it's non-contiguous
'Be careful with the case of contiguous
    n_motor = n_cell \ 2
    Dim outRng As range
    For i = n_motor + 1 To n_cell
        Set outRng = Rg_Union(garageRangeAll.Areas(i), outRng)
    Next
    Set W_GetInsurerCell2 = outRng
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the bottom with premium info
End Function

Function W_GetDealerCell1(ws_name As String)
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the top with age&motor code
    Dim str_arr() As Variant
    'ï¿½ï¿½ï¿½ï¿½á´§, ï¿½ï¿½ï¿½ï¿½, ï¿½ï¿½
    'str_arr = Array("Age", "Year")
    Set garageRangeAll = Rg_FindAllRange("«èÍÁËéÒ§", ws_name)
    'Set garageRangeAll = Rg_FindAllRange("à¸‹à¹ˆà¸­à¸¡à¸«à¹‰à¸²à¸‡", ws_name)
    n_cell = garageRangeAll.Areas.count
'Assume that it has only 2 sets of ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
'Assume that it's non-contiguous
'Be careful with the case of contiguous
    n_motor = n_cell \ 2
    Dim outRng As range
    For i = 1 To n_motor
        Set outRng = Rg_Union(garageRangeAll.Areas(i), outRng)
    Next
    Set W_GetDealerCell1 = outRng
End Function

Function W_GetDealerCell2(ws_name As String)
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the top with age&motor code
    Dim str_arr() As Variant
    'ï¿½ï¿½ï¿½ï¿½á´§, ï¿½ï¿½ï¿½ï¿½, ï¿½ï¿½
    'str_arr = Array("Age", "Year")
    Set garageRangeAll = Rg_FindAllRange("«èÍÁËéÒ§", ws_name)
    'Set garageRangeAll = Rg_FindAllRange("à¸‹à¹ˆà¸­à¸¡à¸«à¹‰à¸²à¸‡", ws_name)
    n_cell = garageRangeAll.Areas.count
'Assume that it has only 2 sets of ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
'Assume that it's non-contiguous
'Be careful with the case of contiguous
    n_motor = n_cell \ 2
    Dim outRng As range
    For i = n_motor + 1 To n_cell
        Set outRng = Rg_Union(garageRangeAll.Areas(i), outRng)
    Next
    Set W_GetDealerCell2 = outRng
'Get ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§ At the bottom with premium info
End Function


Function Pl_Insurer_CarAge(ws_name As String)
    Dim garage_Insurer_Cell, garage_ordered As range
    Set garage_Insurer_Cell = W_GetInsurerCell1(ws_name)
    Set garage_ordered = Rg_ReorderRange(garage_Insurer_Cell)
    Dim outArr() As Variant
    For Each cell In garage_ordered
        carAge = W_GetCarAge(cell.Offset(-1, 0).value)
        outArr = A_Append2(outArr, carAge)
    Next
    Pl_Insurer_CarAge = outArr
    
'Insurer => ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½
'Dealer = > ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§

End Function
Function Pl_Dealer_CarAge(ws_name As String)
    Dim garage_Dealer_Cell, garage_ordered As range
    Set garage_Dealer_Cell = W_GetDealerCell1(ws_name)
    Set garage_ordered = Rg_ReorderRange(garage_Dealer_Cell)
    Dim outArr() As Variant
    For Each cell In garage_ordered
        carAge = W_GetCarAge(cell.Offset(-1, 0).value)
        outArr = A_Append2(outArr, carAge)
    Next
    Pl_Dealer_CarAge = outArr

End Function
Function Pl_Insurer_CoverageINFO(ws_name As String)
'Insurer => ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½
    Dim garageCell, start_num As range
    Set garageCell = W_GetInsurerCell1(ws_name)
    Set garageCell_Reorder = Rg_ReorderRange(garageCell)
    Dim outArr() As Variant
    For Each cell In garageCell_Reorder
        Set start_num = cell.Offset(1, 0)
        Set InfoCell = Sp_SelectFromTL(start_num, 11)
        InfoArr = A_toArray1d(InfoCell)
        outArr = A_Append2(outArr, InfoArr)
    Next
    Pl_Insurer_CoverageINFO = outArr
End Function
Function Pl_Insurer_CoverageINFO2(ws_name As String)
'assume has «èÍÁÍÙè 1 cell
' THAI FIX
    Dim outArr(12) As Variant

'write for 2 cases:  1)Have only 1 line of ¤èÒÃÑ¡ÉÒ¾ÂÒºÒÅ 2) Have 2 lines
    
    'catch_N_MedicalPassenger = Array()
    

End Function
Function Pl2_01_G_TPBI_Person(ws_name As String)
'THAI FIX
    Set ws = Worksheets(ws_name)
    catch_G_TPBIPerson = Array("Í¹ÒÁÑÂºØ¤¤ÅÀÒÂ¹Í¡")
    Set text_cell = Rg_FindAllRange(catch_G_TPBIPerson, ws_name)
End Function
Function Pl2_02_I_TPPD()
'THAI FIX
    catch_I_TPPD = Array("ÃÑº¼Ô´µèÍ·ÃÑ¾ÂìÊÔ¹ºØ¤¤Å")
End Function
Function Pl2_03_J_AccidentDriver()
'THAI FIX
    catch_J_AccidentDriver = Array()
End Function
Function Pl2_04_K_AccidentPassenger()
'THAI FIX
    catch_K_AccidentPassenger = Array()
End Function
Function Pl2_05_L_NoSeats()
'2 cases: 1) Case than has ¤¹ 2)have column of their own
'THAI FIX
    catch_L_NoSeats = Array()
End Function
Function Pl2_06_MN_Medical()
'write for 2 cases:  1)Have only 1 line of ¤èÒÃÑ¡ÉÒ¾ÂÒºÒÅ
'2) Have 2 lines
'THAI FIX
'assume that it has 1 line first
    catch_MN_Medical = Array("¤èÒÃÑ¡ÉÒ¾ÂÒºÒÅ")
End Function
Function Pl2_07_O_BailBond()
'THAI FIX
    catch_O_BailBond = Array()
End Function

Function Pl_Dealer_CoverageINFO(ws_name As String)
'Dealer = > ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
    Dim garageCell, start_num As range
    Set garageCell = W_GetDealerCell1(ws_name)
    Set garageCell_Reorder = Rg_ReorderRange(garageCell)
    Dim outArr() As Variant
    For Each cell In garageCell_Reorder
        Set start_num = cell.Offset(1, 0)
        Set InfoCell = Sp_SelectFromTL(start_num, 11)
        InfoArr = A_toArray1d(InfoCell)
        outArr = A_Append2(outArr, InfoArr)
    Next
    Pl_Dealer_CoverageINFO = outArr
End Function

Function Pl_Insurer_Premium(ws_name As String)
'Insurer => ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½
    Dim ws As Worksheet
    Set ws = Worksheets(ws_name)
    Dim garageCell, start_num As range
    Set garageCell = W_GetInsurerCell2(ws_name)
    Set garageCell_Reorder = Rg_ReorderRange(garageCell)
    
    ws.Select
    Dim outArr() As Variant
    
    For Each cell In garageCell_Reorder
        Set start_num = cell.Offset(2, 0)
        'Set current_region = start_num.CurrentRegion
        'Set current_region = start_num.CurrentRegion.Columns(start_num.column)
        'Set endCell = current_region.SpecialCells(xlCellTypeLastCell)
        'Set premium = range(start_num, endCell)
        Set premium = range(start_num, start_num.End(xlDown))
        InfoArr = A_toArray1d(premium)
        outArr = A_Append2(outArr, InfoArr)
    Next
    Pl_Insurer_Premium = outArr
End Function
Function Pl_Dealer_Premium(ws_name As String)
    Dim garageCell, start_num As range
    Set garageCell = W_GetDealerCell2(ws_name)
    Set garageCell_Reorder = Rg_ReorderRange(garageCell)
    Dim outArr() As Variant
    For Each cell In garageCell_Reorder
        Set start_num = cell.Offset(2, 0)
        Set premium = range(start_num, start_num.End(xlDown))
        InfoArr = A_toArray1d(premium)
        outArr = A_Append2(outArr, InfoArr)
    Next
    Pl_Dealer_Premium = outArr
'Dealer = > ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Ò§
End Function

Function Pl_Insurer_MotorCode(ws_name As String)
    motorCode = Pl_MotorCode(ws_name)
    garage = Pl_Garage(ws_name)
    Dim outArr() As Variant
    For i = LBound(garage) To UBound(garage)
        If garage(i) = "Insurer" Then
            outArr = A_Append2(outArr, motorCode(i))
        End If
    Next
    Pl_Insurer_MotorCode = outArr
End Function

Function Pl_Dealer_MotorCode(ws_name As String)
    motorCode = Pl_MotorCode(ws_name)
    garage = Pl_Garage(ws_name)
    Dim outArr() As Variant
    For i = LBound(garage) To UBound(garage)
        If garage(i) = "Dealer" Then
            outArr = A_Append2(outArr, motorCode(i))
        End If
    Next
    Pl_Dealer_MotorCode = outArr

End Function

Sub W_CreateNeoTemplateRun()
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'TemplatePath = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct\NeoTemplate02.xlsx"
    TemplatePath = "C:\Users\n1603499\OneDrive - Liberty Mutual\Documents\12.02  Rotation2  EastProduct\02.02  AutomateFile Play 04.xlsx"
    newFileName = "Temp01"
    Call W_CreateNeoTemplate(TemplatePath, newFileName)
    
    
End Sub

Sub W_CreateNeoTemplate(TemplatePath, newFileName)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim campaign_name As String
    campaign_name = "Standard1"
    'Open the file
    Dim wb_NeoTemplate As Workbook
    Set wb_NeoTemplate = Workbooks.Open(TemplatePath)
    
    'Save the file as a new file
    Dim newTemplatePath As String
    newTemplatePath = St_RemoveFileName(TemplatePath) & newFileName
    ' Don't show confirmation window
    Application.DisplayAlerts = False
    wb_NeoTemplate.SaveAs fileName:=newFileName
    ' Allow confirmation windows to appear as normal
    Application.DisplayAlerts = True
    
' Info that both use
    VBAFilename = "04.01  VBA Play V04"
    ThisWorkbook.Activate
  ' ################################# Declare variable for extracting data #################################
    'Workbooks("04.01  VBA Play V04").Activate
    both_Make = Pl_Make(campaign_name)
    both_Model = Pl_Model(campaign_name)
    both_SumAssureMin = Pl_SumAssuredMin(campaign_name)
    both_SumAssureMax = Pl_SumAssuredMax(campaign_name)
'Import values for Insurer
'Ins = Insurer
    Ins_coverageINFO = Pl_Insurer_CoverageINFO(campaign_name)
    Ins_MotorCode = Pl_Insurer_MotorCode(campaign_name)
    Ins_CarAge = Pl_Insurer_CarAge(campaign_name)
    Ins_Premium = Pl_Insurer_Premium(campaign_name)
 ' ------------------------------------------ Declare variable for extracting data -----------------------------------------------
 
   ' ################################# Fill Values in Template #################################
    wb_NeoTemplate.Activate
    Dim coverageWS As Worksheet
    Set coverageWS = wb_NeoTemplate.Worksheets("Coverage Input")
    coverageWS.Activate
    Set text_cell01 = coverageWS.usedRange.Find("motorCode")
    Set start_cell01 = text_cell01.Offset(1, 0)
    Set text_cell02 = coverageWS.usedRange.Find("TPBI")
    Set start_cell02 = text_cell02.Offset(1, 0)
    
    Call A_FillValue(Ins_MotorCode, start_cell01, xlDown)
    Call A_FillValue(Ins_coverageINFO, start_cell02, xlToRight)
    
    Set coverageWS = wb_NeoTemplate.Worksheets("Net Premium Input")
    
 ' ------------------------------------------ Fill Values in Template-----------------------------------------------
    

End Sub






