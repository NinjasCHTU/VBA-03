Attribute VB_Name = "Work05_Rotation2"
Function Wb_byUsedRange(xlWorkbook)
'Hard for chatGPT
    Dim ws_arr() As Variant
    Dim usedCellCount() As Variant
    Dim ws_name() As Variant
    Dim inx_arr() As Variant
    Dim out_wsArr() As Variant
    Dim i As Long
    
    ReDim ws_arr(0 To xlWorkbook.Sheets.count - 1)
    ReDim usedCellCount(0 To xlWorkbook.Sheets.count - 1)
    ReDim ws_name(0 To xlWorkbook.Sheets.count - 1)
    ReDim inx_arr(0 To xlWorkbook.Sheets.count - 1)
    ReDim out_wsArr(0 To xlWorkbook.Sheets.count - 1)
    
    For i = 0 To xlWorkbook.Sheets.count - 1
        Set ws_arr(i) = xlWorkbook.Sheets(i + 1)
        inx_arr(i) = i
        'Set temp = ws_arr(i).usedRange
        'Set temp02 = temp.SpecialCells(xlCellTypeConstants)
        'temp03 = temp02.count
        On Error GoTo Err01
        'Error when sheet has only empty cells
        usedCellCount(i) = ws_arr(i).usedRange.SpecialCells(xlCellTypeConstants).count
Back01:
        ws_name(i) = xlWorkbook.Sheets(i + 1).name
        On Error GoTo 0
        
    Next i
    inx_sort = VB_SortBy(inx_arr, usedCellCount, -1)
    'inx_sort = A_ShiftIndex(inx_sort, -1)
    
    For i = LBound(inx_sort) To UBound(inx_sort)
        Set out_wsArr(i - 1) = ws_arr(inx_sort(i))
    Next
    
    Wb_byUsedRange = out_wsArr
    Exit Function
Err01:
    usedCellCount(i) = 0
    GoTo Back01
End Function

Sub Bt_UpdateRedbook()
'Not Done just started !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    Dim filepath As Variant
    Dim wb_redbook As Workbook
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    On Error GoTo Err01
    Set xlApp = New Excel.Application
    On Error GoTo Err02
    'ChDir "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    ChDir "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    On Error GoTo 0
    'Change default path when open the select file window
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    'FileFilter = make the user sees only excel file
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    Set xlApp = New Excel.Application
    If filepath <> False Then
    'UpdateLinks:=0 not to update links for redbookfile
        Set xlWorkbook = xlApp.Workbooks.Open(filepath, UpdateLinks:=0)
    Else
        MsgBox ("The import is canceled")
        Exit Sub
    End If
    Application.ScreenUpdating = True
    On Error GoTo Err01
    ws01 = Wb_byUsedRange(xlWorkbook)
    ' ws_redbook has the most used cells(with text)
    Set ws_redbook = ws01(0)
    makeModelData = W_GetRedBookData(ws_redbook)
    
    MsgBox ("Make Model Import Succesful !!!")
    xlWorkbook.Close
    MsgBox ("File is closed")
    On Error GoTo 0
    Exit Sub
Err01:
    MsgBox ("There's an error during make model import")
    xlWorkbook.Close
    MsgBox ("File is closed")
    Exit Sub
Err02:
    MsgBox ("There's a problem with default path check the VBA code at Bt_UpdateRedbook")
    Exit Sub
End Sub
Function Rg_TopORBottom(rng01, rng02)
    If rng01.row < rng02.row Then
        Rg_TopORBottom = "top"
    ElseIf rng01.row > rng02.row Then
        Rg_TopORBottom = "bottom"
    Else
        Rg_TopORBottom = "same row"
    End If
End Function

Function Rg_LeftORRight(rng01, rng02)
'assume rng01, rng02 to be 1 cell
    If rng01.column < rng02.column Then
        Rg_LeftORRight = "left"
    ElseIf rng01.column > rng02.column Then
        Rg_LeftORRight = "right"
    Else
        Rg_LeftORRight = "same column"
    End If
End Function
Function Rg_PickTopOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are above ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refRow = ref_rng_in.row
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refRow = ref_rng.row
    End If
    
    If include = 1 Then
      refRow = ref_rng.row + 1
    ElseIf include = 0 Then
      refRow = ref_rng.row
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
        curr_row = rng.row
        If curr_row < refRow Then
            Set output_range = Rg_Union(output_range, rng)
        End If
    Next
    
    Set Rg_PickTopOf = output_range
End Function

Function Rg_PickBottomOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are below ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refRow = ref_rng_in.row
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refRow = ref_rng.row
    End If
    
    If include = 1 Then
      refRow = ref_rng.row - 1
    ElseIf include = 0 Then
      refRow = ref_rng.row
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
        curr_row = rng.row
        If curr_row > refRow Then
            Set output_range = Rg_Union(output_range, rng)
        End If
    Next
    Set Rg_PickBottomOf = output_range
End Function

Function Rg_PickLeftOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are left ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refCol = ref_rng_in.column
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refCol = ref_rng.column
    End If
    
    If include = 1 Then
      refCol = ref_rng.column + 1
    ElseIf include = 0 Then
      refCol = ref_rng.column
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
    curr_col = rng.column
    If curr_col < refCol Then
        Set output_range = Rg_Union(output_range, rng)
    End If
    Next
    Set Rg_PickLeftOf = output_range
End Function
Function Rg_PickRightOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are right ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refCol = ref_rng_in.column
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refCol = ref_rng.column
    End If
    
    If include = 1 Then
      refCol = ref_rng.column - 1
    ElseIf include = 0 Then
      refCol = ref_rng.column
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
    curr_col = rng.column
    If curr_col > refCol Then
        Set output_range = Rg_Union(output_range, rng)
    End If
    Next
    Set Rg_PickRightOf = output_range
End Function


