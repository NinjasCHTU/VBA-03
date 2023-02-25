Attribute VB_Name = "Work02_Rotation2_SR"
Sub DuplicateVariablePremiumSheet()
    Dim ws_name As String
    ws_name = "SR1.1"
    v01_motorCode = Pl_MotorCode_Coverage(ws_name, wb)
    v02_Make = Pl_Make(ws_name, wb)
    v03_model = Pl_Model(ws_name, wb)
    v04_age = Pl_CarAge(ws_name, wb)
    v05_SAMin = Pl_SumAssuredMin(ws_name, wb)
    v06_SAMax = Pl_SumAssuredMax(ws_name, wb)
    v07_premium = Pl_Premium(ws_name, wb)
    '********************************************************************** Duplicates Array Variable *****************************************************************************************
    ageGroup = Pl_AgeGroup(ws_name)
    n_models = UBound(v03_model) + 1
    
    w01_motorCode = A_Replicates(v01_motorCode, ageGroup)
    w02_Make = A_Replicates(v02_Make, ageGroup)
    w03_model = A_Replicates(v03_model, ageGroup)
    w04_age = A_Replicates(v04_age, n_models, , , True)
    w05_SAMin = A_Replicates(v05_SAMin, ageGroup)
    w06_SAMax = A_Replicates(v06_SAMax, ageGroup)
    w07_premium = A_Reshape_2dTo1D(v07_premium)
    '------------------------------------------------------------------------------------- Duplicates Array Variable  ---------------------------------------------------------------------
    
End Sub
Function W_GetNum1Line(start_cell)
    Set val_cell = start_cell
    Dim outArr() As Variant
    
    outArr = A_Append2(outArr, start_cell.value)
    Do Until val_cell Is Nothing
        Set val_cell = Rg_NextContainNum(val_cell, xlToRight)
        If val_cell Is Nothing Then
            W_GetNum1Line = outArr
            Exit Function
        End If
        MyVal = St_GetNum(val_cell.value)
        outArr = A_Append2(outArr, MyVal)
    Loop
    W_GetNum1Line = outArr
End Function
Function A_InIfNum(rng)
    Dim outArr() As Variant
    For Each cell In rng
        If IsNumeric(cell.value) And cell.value <> "" Then
            outArr = A_Append(outArr, cell.value)
        End If
    Next cell
    A_InIfNum = outArr
End Function
'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Private Function O_RegKeyRead(i_RegKey As String) As String
    Dim myWS As Object
    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    O_RegKeyRead = myWS.RegRead(i_RegKey)
End Function



' This is needed to get the local path, not the one drive path
Function O_GetDocLocalPath(docPath As String) As String
'return the local path for doc, which is either already a local document or a document on OneDrive
    Const strcOneDrivePart As String = "https://d.docs.live.net/"
    Dim strRetVal As String, bytSlashPos As Byte
    
    strRetVal = docPath & "\"
    If Left(LCase(docPath), Len(strcOneDrivePart)) = strcOneDrivePart Then 'yep, it's the OneDrive path
        'locate and remove the "remote part"
        bytSlashPos = InStr(Len(strcOneDrivePart) + 1, strRetVal, "/")
        strRetVal = Mid(docPath, bytSlashPos)
        'read the "local part" from the registry and concatenate
        strRetVal = O_RegKeyRead("HKEY_CURRENT_USER\Environment\OneDrive") & strRetVal
        strRetVal = Replace(strRetVal, "/", "\") 'slashes in the right direction
        strRetVal = Replace(strRetVal, "%20", " ") 'a space is a space once more
    End If
    O_GetDocLocalPath = strRetVal
    
End Function
Function Wb_GetWB2(Optional defaultFolderPath = "", Optional visiblility = False)
    Dim xlWorkbook As Workbook
    localPath = O_GetDocLocalPath(ThisWorkbook.path)
    folderPath = St_RemoveFileName(localPath)
    If defaultFolderPath = "" Then
        defaultFolderPath = folderPath
    End If
    'Change default path when open the select file window
    ChDir defaultFolderPath
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    
    'FileFilter = make the user sees only excel file
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    If filepath <> False Then
            On Error Resume Next
            Set xlWorkbook = Workbooks(filepath)
            On Error GoTo 0
            'UpdateLinks:=0 not to update links for redbookfile
            If xlWorkbook Is Nothing Then
                Set xlWorkbook = Workbooks.Open(filepath, UpdateLinks:=0)
                'xlWorkbook.visible = visiblility
                Set Wb_GetWB2 = xlWorkbook
            Else
                Set Wb_GetWB2 = xlWorkbook
            End If
    Else
        MsgBox ("The import is canceled")
        Exit Function
    End If
    Application.ScreenUpdating = True
End Function

Function Wb_GetWB(Optional defaultFolderPath = "") As Excel.Workbook
    Dim filepath As Variant
    Dim wb_redbook As Workbook
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    Dim bIsOpen As Boolean
    
    Set xlApp = New Excel.Application
    On Error Resume Next
    
    folderPath = St_RemoveFileName(ThisWorkbook.path)
    If defaultFolderPath = "" Then
        defaultFolderPath = folderPath
    End If
    'Change default path when open the select file window
    ChDir defaultFolderPath
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    
    'FileFilter = make the user sees only excel file
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    
    If filepath <> False Then
            On Error Resume Next
            Set xlWorkbook = GetObject(filepath)
            On Error GoTo 0
            'UpdateLinks:=0 not to update links for redbookfile
            If xlWorkbook Is Nothing Then
                Set xlWorkbook = xlApp.Workbooks.Open(filepath, UpdateLinks:=0)
                Set Wb_GetWB = xlWorkbook
            Else
                Set Wb_GetWB = xlWorkbook
            End If
    Else
        MsgBox ("The import is canceled")
        Exit Function
    End If
    Application.ScreenUpdating = True
End Function


Function OS_GetFilePath()
    Dim filepath As Variant
    Dim wb_redbook As Workbook
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    Set xlApp = New Excel.Application

    On Error GoTo 0
    'Change default path when open the select file window
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    'FileFilter = make the user sees only excel file
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    Set xlApp = New Excel.Application
    If filepath <> False Then
    'UpdateLinks:=0 not to update links for redbookfile
        Set OS_GetFilePath = filepath
    Else
        MsgBox ("The import is canceled")
        Exit Sub
    End If
    Application.ScreenUpdating = True
End Function

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
    Set rng = Selection
    
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
    Set dataRange = Selection
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
    Dim file As Object
    Dim fileName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Dim out_arr() As Variant
    For Each file In folder.Files
        If LCase(Right(file.name, 4)) = ".xls" Or LCase(Right(file.name, 5)) = ".xlsx" Then
            fileName = Left(file.name, (Len(file.name) - 5))
            out_arr = A_Append(out_arr, fileName)
        End If
    Next file
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
        num = St_GetNum(text)
        W_GetCarAge = Array(num, num)
        Exit Function
    End If
    If UBound(numbers) = 1 Then
        num0 = St_GetNum(numbers(0))
        num1 = St_GetNum(numbers(1))
        W_GetCarAge = Array(num0, num1)
    Else
        W_GetCarAge = Array(0, 0)
    End If
End Function


Function St_GetNum(string_in) As Double
    Dim tempString As String
    tempString = ""
    For i = 1 To Len(string_in)
        If IsNumeric(Mid(string_in, i, 1)) Or Mid(string_in, i, 1) = "." Then
            tempString = tempString & Mid(string_in, i, 1)
        End If
    Next i
    If IsNumeric(tempString) Then
        St_GetNum = CDbl(tempString)
    Else
        St_GetNum = 0 ' return 0 if the string does not contain any numeric value
    End If
End Function

Sub A_FillValue(myArray As Variant, start_cell As Variant, Optional direction As XlDirection = xlDown, Optional transpose = False, Optional ws_name = "", Optional overwrite = False, Optional wb = "")
'More general and more powerful than A_printArr
'Done debugging for all cases(Could check more but it seems pretty stable from testing)
'Hard for chatGPT
    Dim rng As range
    Dim mySheet As Worksheet
    Dim row_size, col_size As Integer
    
    On Error GoTo Pass01:
    If wb = "" And TypeName(wb) = "String" Then
        Set wb = ThisWorkbook
    End If
Pass01:
    
    If ws_name = "" Then
        ws_name = ActiveSheet.name
    End If
    On Error GoTo 0
    
    Set mySheet = wb.Worksheets(ws_name)
    
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
        'data = > New 2d array for put it into range
        If direction = xlDown Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
            Set toFill = start_cell.Resize(UBound(myArray) - LBound(myArray) + 1, 1)
            ReDim data(LBound(myArray) To UBound(myArray), 1 To 1)
            For i = LBound(myArray) To UBound(myArray)
                data(i, 1) = myArray(i)
            Next i
            toFill.value = data
            
            'Old Code
            'For i = LBound(myArray) To UBound(myArray)
                'start_cell.Offset(i, 0).value = myArray(i)
            'Next i
            
        ElseIf direction = xlUp Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
            flipedArr = A_Flip(myArray)
            row_size = -UBound(myArray) + LBound(myArray) - 1
            Set toFill = Rg_Resize(start_cell, row_size, 1)
            ReDim data(LBound(myArray) To UBound(myArray), 1 To 1)
            For i = LBound(myArray) To UBound(myArray)
                data(i, 1) = flipedArr(i)
            Next i
            toFill.value = data
        ElseIf direction = xlToRight Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
            
            col_size = UBound(myArray) - LBound(myArray) + 1
            Set toFill = Rg_Resize(start_cell, 1, col_size)
            ReDim data(1 To 1, LBound(myArray) To UBound(myArray))
            For i = LBound(myArray) To UBound(myArray)
                data(1, i) = myArray(i)
            Next i
            toFill.value = data
        ElseIf direction = xlToLeft Then
            Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
            flipedArr = A_Flip(myArray)
            col_size = -UBound(myArray) + LBound(myArray) - 1
            Set toFill = Rg_Resize(start_cell, 1, col_size)
            ReDim data(1 To 1, LBound(myArray) To UBound(myArray))
            For i = LBound(myArray) To UBound(myArray)
                data(1, i) = flipedArr(i)
            Next i
            toFill.value = data
        Else
            MsgBox "Invalid direction"
        End If
'For 2d Array case
    ElseIf A_NDim(myArray) = 2 Then
        Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
        myArray2 = IIf(transpose, VB_Transpose(myArray), myArray)
        n_row = UBound(myArray2, 1) - LBound(myArray2, 1) + 1
        n_col = UBound(myArray2, 2) - LBound(myArray2, 1) + 1
        
        
        Set toFill = Rg_Resize(start_cell, n_row, n_col)
        toFill.value = myArray2
        
        
'         If direction = xlDown Then
'             If direction2 = "" Or direction2 = xlToRight Then
'             'assume that index start with 0 otherwise it won't work
'             'Because Offset needs 0 in order for it to select that start_cell
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(j, i).value = myArray(i, j)
'                     Next j
'                 Next i
'             ElseIf direction2 = xlToLeft Then
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToLeft))
                
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(j, -i).value = myArray(i, j)
'                     Next j
'                 Next i
'             Else
'                 MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
'             End If
'         ElseIf direction = xlUp Then
'             If direction2 = "" Or direction2 = xlToRight Then
'             'assume that index start with 0 otherwise it won't work
'             'Because Offset needs 0 in order for it to select that start_cell
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToRight))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(-j, i).value = myArray(i, j)
'                     Next j
'                 Next i
'             ElseIf direction2 = xlToLeft Then
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlToLeft))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(-j, -i).value = myArray(i, j)
'                     Next j
'                 Next i
'             Else
'                 MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
'             End If
'         ElseIf direction = xlToRight Then
'             If direction2 = "" Or direction2 = xlDown Then
'             'assume that index start with 0 otherwise it won't work
'             'Because Offset needs 0 in order for it to select that start_cell
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(i, j).value = myArray(i, j)
'                     Next j
'                 Next i
'             ElseIf direction2 = xlUp Then
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlUp))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(-i, j).value = myArray(i, j)
'                     Next j
'                 Next i
'             Else
'                 MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
'             End If
'         ElseIf direction = xlToLeft Then
'             If direction2 = "" Or direction2 = xlDown Then
'             'assume that index start with 0 otherwise it won't work
'             'Because Offset needs 0 in order for it to select that start_cell
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlDown))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(i, -j).value = myArray(i, j)
'                     Next j
'                 Next i
'             ElseIf direction2 = xlUp Then
'                 Set start_cell = IIf(start_cell.value = "" Or overwrite, start_cell, Rg_NextNoTextCell(start_cell, xlUp))
'                 For i = 0 To UBound(myArray, 1)
'                     For j = 0 To UBound(myArray, 2)
'                         start_cell.Offset(-i, -j).value = myArray(i, j)
'                     Next j
'                 Next i
'             Else
'                 MsgBox ("Invalid direction2: Please use xlUp,xlDown,xlToRight,xlToLeft, or blank")
'             End If
'         Else
'             MsgBox "Invalid direction"
'         End If
    
'     Else
'         MsgBox ("Not support array with dimesion >2")
'     End If

' End Sub
' 'thisWBName = "04.01  VBA Play V04"

' Function A_FindFromHook(search_list, Optional offset_row = 1, Optional offset_col = 0, Optional ws_name)
' 'search_list could be string or Array
' ' If there are many cells that have the same word it will get only the 1st one
'     Dim outArr As Variant
'     If IsArray(search_list) Then
'         For i = LBound(search_list) To UBound(search_list)
'             curr_str = search_list(i)
'             foundArr = A_FindFromHookH1(curr_str, offset_row, offset_col, ws_name)
'             outArr = A_Extend(outArr, foundArr)
'         Next
'     Else
'         outArr = A_FindFromHookH1(search_list, offset_row, offset_col, ws_name)
'     End If
'     A_FindFromHook = outArr
    End If
End Sub
'thisWBName = "04.01  VBA Play V04"

Function A_FindFromHook(search_list, Optional offset_row = 1, Optional offset_col = 0, Optional ws = "", Optional wb = "")
'search_list could be string or Array
' If there are many cells that have the same word it will get only the 1st one

    ws_name = Ws_WS_at_WB(ws, wb, False)
    Dim outArr As Variant
    If IsArray(search_list) Then
        For i = LBound(search_list) To UBound(search_list)
            curr_str = search_list(i)
            foundArr = A_FindFromHookH1(curr_str, offset_row, offset_col, ws_name, wb)
            outArr = A_Extend(outArr, foundArr)
        Next
    Else
        outArr = A_FindFromHookH1(search_list, offset_row, offset_col, ws_name, wb)
    End If
    A_FindFromHook = outArr
End Function

Function A_FindFromHookH1(search_str, Optional offset_row = 1, Optional offset_col = 0, Optional ws = "", Optional wb = "")
    Dim ans_arr() As Variant
    
    Set ws02 = Ws_WS_at_WB(ws, wb)

    Dim SearchRange As range
    Set SearchRange = ws02.usedRange
    
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
Function Pl3_ProductCode(ws, Optional wb = "", Optional docCase = "case01_Easy")
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catchArr = Array("Product code")
    If docCase = "case01_Easy" Then
        Set text_cell = Rg_FindAllRange(catchArr, ws, xlPart, wb)
        Set val_cell = text_cell.Offset(0, 1)
        res = val_cell.value
    Else
    End If
    Pl3_ProductCode = res
End Function
Function Pl3_SMCode(ws, Optional wb = "", Optional docCase = "case01_Easy")
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catchArr = Array("SM CODE")
    If docCase = "case01_Easy" Then
        Set text_cell = Rg_FindAllRange(catchArr, ws, xlPart, wb)
        text = text_cell.Offset(0, 1).value
        res = St_GetBefore(text, " ")
    Else
    End If
    Pl3_SMCode = res
End Function
Function Pl3_CreateFileName(ws, Optional wb = "", Optional docCase = "case01_Easy")

    If docCase = "case01_Easy" Then
        productCode = Pl3_ProductCode(ws, wb, docCase)
        SMCode = Pl3_SMCode(ws, wb, docCase)
        outStr = productCode & "_" & SMCode
    Else
    End If
    Pl3_CreateFileName = outStr
End Function

Function Pl_Make(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catchArr = Array("Make", "ÂÕèËéÍ")
    If docCase = "case01_Easy" Then
        Set text_cell = Rg_FindAllRange(catchArr, ws, xlPart, wb)
        Set start = text_cell.Offset(1, 0)
        Set val_cell = Rg_PickTilEnd(start, xlDown)
        outArr = A_RepMerge(val_cell)
    Else
    End If
    Pl_Make = outArr
    
'Pl = Pull data
End Function

Function Pl_Model(ws, Optional wb = "", Optional docCase = "case01_Easy")
    catchArr = Array("Model", "ÃØè¹")
    ws_name = Ws_WS_at_WB(ws, wb, False)
    If docCase = "case01_Easy" Then
        Set text_cell = Rg_FindAllRange(catchArr, ws, xlPart, wb)
        Set start = text_cell.Offset(1, 0)
        Set val_cell = Rg_PickTilEnd(start, xlDown)
        outArr = A_RepMerge(val_cell)
    Else
    End If
    Pl_Model = outArr
End Function
Function Pl_SumAssuredMin(ws, Optional wb = "", Optional docCase = "case01_Easy")
' à¸—à¸¸à¸™à¸›à¸£à¸°à¸à¸±à¸™ => ï¿½Ø¹ï¿½ï¿½Ð¡Ñ¹
    If docCase = "case01_Easy" Then
        catchArr = Array("·Ø¹»ÃÐ¡Ñ¹")
        ws_name = Ws_WS_at_WB(ws, wb, False)
        Set text_cell = Rg_FindSomeRanges(catchArr, -1, ws_name, , xlWhole, wb)
        Set value_cell = Rg_PickTilEnd(text_cell.Offset(1, 0), xlDown)
        val_arr = A_toArray1d(value_cell)
        Pl_SumAssuredMin = val_arr
    Else
    End If
End Function
Function Pl_SumAssuredMax(ws, Optional wb = "", Optional docCase = "case01_Easy")
    If docCase = "case01_Easy" Then
        catchArr = Array("·Ø¹»ÃÐ¡Ñ¹")
        ws_name = Ws_WS_at_WB(ws, wb, False)
        Set text_cell = Rg_FindSomeRanges(catchArr, -1, ws_name, , xlWhole, wb)
        Set value_min = text_cell.Offset(1, 0)
        Set value_max = Rg_PickTilEnd(value_min.Offset(0, 2), xlDown)
        'Set value_cell = Rg_PickTilEnd(text_cell.Offset(1, 2), xlDown)
        val_arr = A_toArray1d(value_max)
        Pl_SumAssuredMax = val_arr
    Else
    End If
End Function
Function Pl_MotorCode_MakeModel(ws, Optional wb = "", Optional docCase = "case01_Easy")
    ws_name = Ws_WS_at_WB(ws, wb, False)
    If docCase = "case01_Easy" Then
        catchArr = Array("ÃËÑÊ")
        Set text_cell = Rg_FindSomeRanges(catchArr, -1, ws_name, , , wb)
        Set code_cell = Rg_PickTilEnd(text_cell.Offset(1, 0), xlDown)
        code_arr = A_toArray1d(code_cell)
        Pl_MotorCode_MakeModel = code_arr
    Else
    End If
    
End Function
Function Pl_MotorCode_Coverage(ws, Optional wb, Optional docCase = "case01_Easy")
'à¸£à¸«à¸±à¸ª => ï¿½ï¿½ï¿½ï¿½
    Dim SearchRange, text_cell As range
    
    If docCase = "case01_Easy" Then
        Set ws02 = Ws_WS_at_WB(ws, wb)
        'search_str = "à¸£à¸«à¸±à¸ª"
        catchArr = Array("ÃËÑÊ")
        'Set SearchRange = ws.usedRange
        Set text_cell = Rg_FindAllRangeH1(catchArr, ws02, , wb)
        Dim outArr() As Variant
        Dim myText As String
        For Each cell In text_cell
            myText = cell.value
            motorCode = St_GetNum(myText)
            If Not A_isInArr(outArr, motorCode) And motorCode <> 0 Then
                outArr = A_Append2(outArr, motorCode)
            End If
            
        Next cell
    Else
    End If
    Pl_MotorCode_Coverage = outArr
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
Function Pl_AVType(ws, Optional wb = "", Optional docCase = "case01_Easy")
    Set ws_convert = Ws_WS_at_WB(ws, wb)
    ws_name = Ws_WS_at_WB(ws, wb, False)
    If docCase = "case01_Easy" Then
        catchArr = Array("»ÃÐàÀ·")
        Set text_cell = Rg_FindAllRange(catchArr, ws_name, , wb)
        num = St_GetNum(text_cell.value)
        If num = 1 Then
            outString = "AV1"
        ElseIf num = 2 Then
            outString = "AV2"
        Else
            outString = "AV5"
        End If
    Else
    End If
    Pl_AVType = outString
End Function
Function Pl_CarAge(ws, Optional wb = "", Optional docCase = "case01_Easy")
    Set ws_convert = Ws_WS_at_WB(ws, wb)
    ws_name = Ws_WS_at_WB(ws, wb, False)
    Dim outArr() As Variant
    If docCase = "case01_Easy" Then
        catchArr = Array("àºÕéÂÊØ·¸Ô")
        Set text_cell = Rg_FindAllRange(catchArr, ws_name, xlWhole, wb)
        For Each cell In text_cell
            Set val_cell = cell.Offset(-1, 0)
            carAge = W_GetCarAge(val_cell.value)
            outArr = A_Append2(outArr, carAge)
        Next
    Else
    End If
    Pl_CarAge = outArr
End Function
Function Pl_Premium(ws, Optional wb = "", Optional docCase = "case01_Easy")
    Set ws_convert = Ws_WS_at_WB(ws, wb)
    ws_name = Ws_WS_at_WB(ws, wb, False)
    Dim outArr() As Variant
    If docCase = "case01_Easy" Then
        catchArr = Array("àºÕéÂÊØ·¸Ô")
        Set text_cell = Rg_FindAllRange(catchArr, ws_name, xlWhole, wb)
        For Each cell In text_cell
            Set start_cell = cell.Offset(1, 0)
            Set val_cell = Rg_PickTilEnd(start_cell, xlDown)
            temp = A_toArray1d(val_cell)
            outArr = A_HStack(outArr, temp)
        Next
    Else
    End If
    Pl_Premium = outArr
End Function
Function Pl_AgeGroup(ws, Optional wb = "", Optional docCase = "case01_Easy")
    Set ws_convert = Ws_WS_at_WB(ws, wb)
    ws_name = Ws_WS_at_WB(ws, wb, False)
    If docCase = "case01_Easy" Then
        catchArr = Array("àºÕéÂÊØ·¸Ô")
        Set text_cell = Rg_FindAllRange(catchArr, ws_name, xlWhole, wb)
        age_group = text_cell.count
    Else
    End If
    Pl_AgeGroup = age_group
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
'******************************************************* Cases String ****************************************************************************
'DocCase
'"case01_Easy", "case02_MoreGarage"
'GarageCase = "Insurer", "Dealer"
Function Pl_CoverageINFO(ws, Optional wb = "", Optional docCase = "case01_Easy")
    Dim outArr(), v06() As Variant
    ws_name = Ws_WS_at_WB(ws, wb, False)
    If docCase = "case01_Easy" Then
        v01 = Pl2_01_G_TPBI_Person(ws_name, wb, docCase)
        v02 = Pl2_02_I_TPPD(ws_name, wb, docCase)
        v03 = Pl2_03_J_AccidentDriver(ws_name, wb, docCase)
        v04 = Pl2_04_K_AccidentPassenger(ws_name, wb, docCase)
        v05 = Pl2_05_L_NoSeats(ws_name, wb, docCase)
        v06 = Pl2_06_MN_Medical(ws_name, wb, docCase)
        v07 = Pl2_07_O_BailBond(ws_name, wb, docCase)
        outArr = v01
        outArr = A_Append2(outArr, v02)
        outArr = A_Append2(outArr, v03)
        outArr = A_Append2(outArr, v04)
        outArr = A_Append2(outArr, v05)
        outArr = A_Extend(outArr, v06)
        outArr = A_Append2(outArr, v07)
    
    End If
    Pl_CoverageINFO = outArr
    
End Function
'***************************************************************** Coverage INFO  *****************************************************************
Function Pl2_01_G_TPBI_Person(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
'Done
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("Í¹ÒÁÑÂºØ¤¤ÅÀÒÂ¹Í¡")
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    Set val_cell = text_cell
    Dim outArr() As Variant
    If docCase = "case01_Easy" Then
        
        Set val_cell01 = Rg_NextContainNum(val_cell, xlToRight)
        Set val_cell02 = val_cell01.Offset(1, 0)
        arr01 = W_GetNum1Line(val_cell01)
        arr02 = W_GetNum1Line(val_cell02)
        outArr = A_Append2(outArr, arr01)
        outArr = A_Append2(outArr, arr02)
    End If
    Pl2_01_G_TPBI_Person = outArr
End Function
Function Pl2_02_I_TPPD(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
'Done
    catch_by = Array("ÃÑº¼Ô´µèÍ·ÃÑ¾ÂìÊÔ¹ºØ¤¤Å")
    ws_name = Ws_WS_at_WB(ws, wb, False)
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    If docCase = "case01_Easy" Then
        Set start_cell = Rg_NextContainNum(text_cell, xlToRight)
        'outArr = W_GetNum1Line(val_cell)
        Set val_cell = start_cell
        Dim outArr() As Variant
        
        outArr = A_Append2(outArr, start_cell.value)
        Do Until val_cell Is Nothing
            Set val_cell = Rg_NextContainNum(val_cell, xlToRight)
            If val_cell Is Nothing Then
                Pl2_02_I_TPPD = outArr
                Exit Function
            End If
            MyVal = St_GetNum(val_cell.value)
            outArr = A_Append2(outArr, MyVal)
        Loop
        Pl2_02_I_TPPD = outArr
    End If
    Pl2_02_I_TPPD = outArr
End Function
Function Pl2_03_J_AccidentDriver(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
'Not Done I don't know what's circular formula is all about
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("Ã.Â. 01")
    Dim outArr() As Variant
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    
    If docCase = "case01_Easy" Then
        Set text_cell02 = text_cell.Offset(2, 0)
        Set val_cell = Rg_NextContainNum(text_cell02, xlToRight)
'Not Done I don't know what's circular formula is all about
        'arr02 = W_GetNum1Line(val_cell)
        outArr = W_GetNum1Line(val_cell)
    End If

    Pl2_03_J_AccidentDriver = outArr
    
End Function
Function Pl2_04_K_AccidentPassenger(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("Ã.Â. 01")
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)

    If docCase = "case01_Easy" Then
        Set text_cell02 = text_cell.Offset(3, 0)
        Set val_cell = Rg_NextContainNum(text_cell02, xlToRight)
        outArr = W_GetNum1Line(val_cell)
    End If
    Pl2_04_K_AccidentPassenger = outArr
End Function
Function Pl2_05_L_NoSeats(ws, Optional wb = "", Optional docCase = "case01_Easy")
'2 cases: 1) Case than has ¤¹ 2)have column of their own
'THAI FIX
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("¨Ó¹Ç¹¼Ùé¢Ñº¢Õè")
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    If docCase = "case01_Easy" Then
        Set val_cell = Rg_NextContainNum(text_cell, xlToRight)
        outArr = W_GetNum1Line(val_cell)
    End If
    Pl2_05_L_NoSeats = outArr
End Function
Function Pl2_06_MN_Medical(ws, Optional wb = "", Optional docCase = "case01_Easy")
'write for 2 cases:  1)Have only 1 line of ¤èÒÃÑ¡ÉÒ¾ÂÒºÒÅ
'2) Have 2 lines
'THAI FIX
'assume that it has 1 line first
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("Ã.Â. 02")
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    Dim outArr() As Variant
    If docCase = "case01_Easy" Then
        Set text_cell02 = text_cell.Offset(1, 0)
        Set val_cell_driver = Rg_NextContainNum(text_cell02, xlToRight)
        outArrDriver = W_GetNum1Line(val_cell_driver)
        
        Set text_cell03 = text_cell.Offset(3, 0)
        Set val_cell_driver = Rg_NextContainNum(text_cell03, xlToRight)
        
        outArrPassenger = W_GetNum1Line(val_cell_driver)
        outArr = A_Append2(outArr, outArrDriver)
        outArr = A_Append2(outArr, outArrPassenger)
    End If
    Pl2_06_MN_Medical = outArr
    
End Function
Function Pl2_07_O_BailBond(ws, Optional wb = "", Optional docCase = "case01_Easy")
'THAI FIX
    ws_name = Ws_WS_at_WB(ws, wb, False)
    catch_by = Array("Ã.Â. 03")
    Set text_cell = Rg_FindAllRange(catch_by, ws_name, , wb)
    If docCase = "case01_Easy" Then
        Set val_cell = Rg_NextContainNum(text_cell, xlToRight)
        outArr = W_GetNum1Line(val_cell)
    End If
    Pl2_07_O_BailBond = outArr
End Function
'--------------------------------------------------------------------------Coverage INFO ---------------------------------------------------------------------------------------------------------
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
Sub W_CreateNeoTemplate1_1_Run()
    Call W_CreateNeoTemplate1_1
End Sub
Sub W_DeclareVariableTest01()
    
End Sub
Sub W_CreateNeoTemplate1_1()
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim wb_NeoTemplate As Workbook
    Dim campaign_ws_name As String
    
    pathSheet = "·ÕèÍÂÙèä¿Åì"
    
    newFileName = Rg_FindAllRange("New File Name", pathSheet, xlPart).Offset(0, 1).value
    folder = Rg_FindAllRange("Template Folder", pathSheet, xlPart).Offset(0, 1).value
    fileName = Rg_FindAllRange("Template File Name", pathSheet, xlPart).Offset(0, 1).value
    TemplatePath = folder & "\" & fileName
    outputFolder = Rg_FindAllRange("Output", pathSheet, xlPart).Offset(0, 1).value
    defaultFolder = Rg_FindAllRange("àÅ×Í¡µÒÃÒ§", pathSheet, xlPart).Offset(0, 1).value
    
    'Set wb_NeoTemplate = Wb_GetWB3(, TemplatePath)
    Set wb_NeoTemplate = Wb_GetWB4(, TemplatePath)
    Set wb_PremTable = Wb_GetWB3(defaultFolder)

    campaign_ws_name = wb_PremTable.Sheets(1).name
    'Open the file
    
    'outputFolder = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    
    
    
    '*************************************************************** Different ways to refer/open to wb_NeoTemplate  ***************************************************************'
    'Set wb_NeoTemplate = Wb_GetWB(defaultFolder)
    'Set wb_NeoTemplate = ThisWorkbook
    
    
    'Set wb_NeoTemplate = Workbooks("NeoTemplate 04_ForVBA.xlsx")
    
    'Set wb_NeoTemplate = Workbooks.Open()
    '*************************************************************** Different ways to refer/open to wb_NeoTemplate  ***************************************************************'
    
    'Save the file as a new file
    Dim newTemplatePath As String
    newTemplatePath = outputFolder & "\" & newFileName
    newTemplatePath2 = newTemplatePath & ".xlsx"
    If Dir(newTemplatePath2) = "" Then
'What if the file is already opened?
        'MsgBox "File not found."
    Else
'        On Error GoTo Err01
'        Set old_file = Workbooks(newFileName)
'        old_file.Close False
'        On Error GoTo 0
'        On Error GoTo -1
Err01:
        Kill newTemplatePath2
    End If
    ' Don't show confirmation window
    Application.DisplayAlerts = False
    wb_NeoTemplate.SaveAs fileName:=newTemplatePath
    ' Allow confirmation windows to appear as normal
    Application.DisplayAlerts = True
    
' Info that both use
    VBAFilename = "04.01  VBA Play V04"
    ThisWorkbook.Activate
    docCase = "case01_Easy"
'DocCase: "case01_Easy"



  ' ################################# Coverage Variable #################################
    v_motorCode = Pl_MotorCode_Coverage(campaign_ws_name, wb_PremTable, docCase)
    v_coverageINFO = Pl_CoverageINFO(campaign_ws_name, wb_PremTable, docCase)
' ------------------------------------------ Coverage Variable  -----------------------------------------------
  ' ################################# Net Premium Input Variable #################################
    'Workbooks("04.01  VBA Play V04").Activate
    v01_motorCode = Pl_MotorCode_MakeModel(campaign_ws_name, wb_PremTable, docCase)
    v02_Make = Pl_Make(campaign_ws_name, wb_PremTable, docCase)
    v03_model = Pl_Model(campaign_ws_name, wb_PremTable, docCase)
    v04_age = Pl_CarAge(campaign_ws_name, wb_PremTable, docCase)
    v05_SAMin = Pl_SumAssuredMin(campaign_ws_name, wb_PremTable, docCase)
    v06_SAMax = Pl_SumAssuredMax(campaign_ws_name, wb_PremTable, docCase)
    v07_premium = Pl_Premium(campaign_ws_name, wb_PremTable, docCase)
 ' ------------------------------------------ Net Premium Input Variable  -----------------------------------------------
    '********************************************************************** Duplicates Array Variable *****************************************************************************************
    ageGroup = Pl_AgeGroup(campaign_ws_name, wb_PremTable, docCase)
    n_models = UBound(v03_model) + 1
    av = Pl_AVType(campaign_ws_name, wb_PremTable, docCase)
    
    w01_motorCode = A_Replicates(v01_motorCode, ageGroup)
    w02_Make = A_Replicates(v02_Make, ageGroup)
    w03_model = A_Replicates(v03_model, ageGroup)
    w04_age = A_Replicates(v04_age, n_models, , , True)
    w05_SAMin = A_Replicates(v05_SAMin, ageGroup)
    w06_SAMax = A_Replicates(v06_SAMax, ageGroup)
    w07_premium = A_Reshape_2dTo1D(v07_premium)
    '------------------------------------------------------------------------------------- Duplicates Array Variable  ---------------------------------------------------------------------
    

    wb_NeoTemplate.Activate
    Dim coverageWS As Worksheet
    'There will be a problem when they change sheet name
    Set coverageWS = wb_NeoTemplate.Worksheets("Coverage Input")
    Set PremiumWS = wb_NeoTemplate.Sheets("Net Premium Input")
    coverageWS.Activate
   ' ################################# Find text cell in "CoverageInput" #################################
    On Error Resume Next
    'Set text_motorCode = coverageWS.usedRange.Find("motorCode")
    'Set start_motorCode = text_cell01.Offset(1, 0)
    Set start_motorCode = coverageWS.range("E2")
    'Set text_Coverage = coverageWS.usedRange.Find("TPBI")
    'Set start_Coverage = text_cell02.Offset(1, 0)
    Set start_Coverage = coverageWS.range("G2")
    On Error GoTo 0
         ' ################################# Fill Values in "Coverage Input" #################################
    Call A_FillValue(v_motorCode, start_motorCode, xlDown, True, "Coverage Input", , wb_NeoTemplate)
    Call A_FillValue(v_coverageINFO, start_Coverage, xlToRight, True, "Coverage Input", , wb_NeoTemplate)
    
     ' ------------------------------------------ Fill Values in "Coverage Input"  -----------------------------------------------

     
        ' ################################# Find text cell in "Net Premium Input" #################################
    Set text_I = Rg_FindAllRange("motorCode", PremiumWS, xlPart, wb_NeoTemplate)
    Set start_I_MotorCode = text_I.Offset(1, 0)
    'Set start_I_MotorCode = text_I.Offset(1, 0)
    Set text_J = Rg_FindAllRange("vehicleMake", PremiumWS, xlPart, wb_NeoTemplate)
    Set start_J_Make = text_J.Offset(1, 0)
    'Set start_J_Make = PremiumWS.range("J2")
    Set text_L = Rg_FindAllRange("vehicleModel", PremiumWS, xlPart, wb_NeoTemplate)
    Set start_L_Model = text_L.Offset(1, 0)
    'Set start_L_Model = PremiumWS.range("L2")
    
    Set text_R = Rg_FindAllRange("vehicleAge", PremiumWS, xlPart, wb_NeoTemplate)
    Set start_R_Age = text_R.Offset(1, 0)
    
    If av = "AV1" Then
        Set text_SA = Rg_FindAllRange("av1", PremiumWS, xlPart, wb_NeoTemplate)
        Set start_SA_Min = text_SA.Offset(1, 0)
        Set start_SA_Max = text_SA.Offset(1, 1)
    End If
    
    Set text_AT = Rg_FindAllRange("Net Premium", PremiumWS, xlPart, wb_NeoTemplate)
    Set start_AT_Premium = text_AT.Offset(1, 0)
    
     ' ------------------------------------------ Find text cell in "Net Premium Input"  -----------------------------------------------
     ' ################################# Fill Values in "Net Premium Input" #################################
     
    Call A_FillValue(w01_motorCode, start_I_MotorCode, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w02_Make, start_J_Make, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w03_model, start_L_Model, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w04_age, start_R_Age, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w05_SAMin, start_SA_Min, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w06_SAMax, start_SA_Max, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    Call A_FillValue(w07_premium, start_AT_Premium, xlDown, , "Net Premium Input", , wb_NeoTemplate)
    
    Application.DisplayAlerts = False
    wb_NeoTemplate.Save
    wb_NeoTemplate.Close
    wb_PremTable.Close
    Application.DisplayAlerts = True
    MsgBox ("Neo Template is successfully created !!!")
     ' ------------------------------------------ Fill Values in "Net Premium Input"  -----------------------------------------------
 ' ------------------------------------------ Fill Values in Template-----------------------------------------------
    

End Sub





