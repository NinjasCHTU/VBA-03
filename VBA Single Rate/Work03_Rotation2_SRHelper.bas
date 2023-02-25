Attribute VB_Name = "Work03_Rotation2_SRHelper"
Function Ws_WS_at_WB(ws, wb, Optional outputOption = True)
'If outputWS = True => output WS
'If outputWS = False => ws_name as string
'wb could be workbook or string or missing
'ws could be worksheet or string or missing
    On Error GoTo Pass01:
    If IsMissing(wb) Then
        Set wb = ThisWorkbook
    ElseIf TypeName(wb) = "Range" Then
        Set wb = Workbooks(wb.value)
    ElseIf wb = "" Then
        Set wb = ThisWorkbook
    Else
    'wb is string
        Set wb = Workbooks(wb)
Pass01:
    End If
    On Error GoTo 0
    On Error GoTo -1
    
    On Error GoTo Pass02:
    wb_name = wb.name
    If IsMissing(ws) Then
        Set ws02 = ThisWorksheet
    ElseIf TypeName(ws) = "Range" Then
        Set ws02 = wb.Worksheets(ws.value)
    ElseIf TypeName(ws) <> "String" Then
        Set ws02 = ws
    ElseIf ws = "" Then
        Set ws02 = ThisWorksheet
    Else
    'ws is string
        Set ws02 = wb.Worksheets(ws)
    End If
    On Error GoTo 0
    On Error GoTo -1
Pass02:
    Set outputWS = wb.Worksheets(ws02.name)
    
    
    
    If outputOption Then
        Set Ws_WS_at_WB = outputWS
    Else
        Ws_WS_at_WB = outputWS.name
    End If
    
End Function
Function Wb_GetWB4(Optional defaultPath = "", Optional filepath)
'Can use direct filepath to open as well
    'It opens in the background and doesn't error when file is already opened
    If defaultPath <> "" Then
        ChDir defaultPath
    End If
    
    Dim wb_used As range
    Application.ScreenUpdating = False
    If IsMissing(filepath) Then
        filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    End If
    If filepath <> False Then
        On Error GoTo Err01
        'Set wb01 = GetObject(filepath)
        Set wb01 = Application.Workbooks.Open(filepath)
    End If
    Set Wb_GetWB4 = wb01
    Application.ScreenUpdating = True
    Exit Function
Err01:
    fileName = St_FileNameFromPath(filepath)
    Set wb01 = Workbooks(fileName)
    Set Wb_GetWB3 = wb01
End Function
Function Wb_GetWB3(Optional defaultPath = "", Optional filepath)
'Can use direct filepath to open as well
    'It opens in the background and doesn't error when file is already opened
    If defaultPath <> "" Then
        ChDir defaultPath
    End If
    
    Dim wb_used As range
    Application.ScreenUpdating = False
    If IsMissing(filepath) Then
        filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    End If
    If filepath <> False Then
        On Error GoTo Err01
        Set wb01 = GetObject(filepath)
        'Set wb01 = Application.Workbooks.Open(filepath)
    End If
    Set Wb_GetWB3 = wb01
    Application.ScreenUpdating = True
    Exit Function
Err01:
    fileName = St_FileNameFromPath(filepath)
    Set wb01 = Workbooks(fileName)
    Set Wb_GetWB3 = wb01
End Function
Function A_HStackH1(arr1, arr2)
    ' arr1 and arr2 is 2d Array
    'What if arr2 is 1d array
    
    n_row1 = UBound(arr1, 1)
    n_col1 = UBound(arr1, 2)
    n_row2 = UBound(arr2, 1)
    n_col2 = UBound(arr2, 2)
    If n_row1 <> n_row2 Then
        A_HStackH1 = "The number of column is not equal"
        Exit Function
    End If
    Dim out_arr() As Variant
    ReDim out_arr(n_row1, n_col1 + n_col2 + 1)
    For i = 0 To n_row1
        For j = 0 To n_col1
            out_arr(i, j) = arr1(i, j)
        Next
    Next
    For i = 0 To n_row2
        For j = 0 To n_col2
            out_arr(i, j + n_col1 + 1) = arr2(i, j)
        Next
    Next
    A_HStackH1 = out_arr

End Function
Function A_HStack(arr1, arr2)
'Upgrade: Right now it supports both 1d 2d and empty
'But A_VStack is not yet
'Done A_HStack that works for both 1d and 2d
'In case one of them is empty it will return the non-empty same d another array(if input is 1d output from this is also 1d)
    
    If A_NDim(arr1) = 1 Then
        arr1_2d = A_1dTo2d(arr1)
    ElseIf A_IsEmpty(arr1) Then
        A_HStack = arr2
        Exit Function
    Else
        arr1_2d = arr1
    End If
    
    If A_NDim(arr2) = 1 Then
        arr2_2d = A_1dTo2d(arr2)
    ElseIf A_IsEmpty(arr2) Then
        A_HStack = arr1
        Exit Function
    Else
        arr2_2d = arr2
    End If
    A_HStack = A_HStackH1(arr1_2d, arr2_2d)
End Function
Function A_GetRow(arr As Variant, i) As Variant
    Dim numCols As Long
    Dim outputArr() As Variant
    Dim j As Long
    
    numCols = UBound(arr, 2) - LBound(arr, 2) + 1
    ReDim outputArr(0 To numCols - 1)
    
    For j = LBound(arr, 2) To UBound(arr, 2)
        outputArr(j) = arr(i, j)
    Next j
    
    A_GetRow = outputArr
End Function

Function A_GetColumn(arr As Variant, i) As Variant
    Dim numRows As Long
    Dim columnArr() As Variant
    Dim j As Long
    
    numRows = UBound(arr, 1) - LBound(arr, 1) + 1
    
    ReDim columnArr(0 To numRows - 1)
    
    For j = 0 To numRows - 1
        columnArr(j) = arr(j, i)
    Next j
    
    A_GetColumn = columnArr
End Function

Function A_HStackRep(arr_in, n)
    temp_arr = arr_in
    For i = 1 To n - 1
        temp_arr = A_HStackH1(temp_arr, arr_in)
    Next i
    A_HStackRep = temp_arr

End Function
Function A_Replicates(arr, n, Optional copyDirection As XlDirection = xlDown, Optional byRow = True, Optional same_together = False)
'same together = old elments are near each other when coppied
'byRow = True copy by row
'byRow = True copy by col(little confusing see example)
'xlUp or xlDown perform the same way
'xlToRight or xlToLeft perform the same way
'************************************************************************ Grand Function
    Dim temp(), outArr(), block(), row(), line() As Variant
    n_dim = A_NDim(arr)
    Dim i As Long
    If n_dim = 1 Then
        If same_together Then
            outArr = A_VRepArr(arr, n)
        Else
        'Normal oneline
            For j = 0 To n - 1
                outArr = A_Extend(outArr, arr)
            Next
            
        End If
        
    ElseIf n_dim = 2 Then
        If same_together Then
            If byRow Then
                If copyDirection = xlDown Or copyDirection = xlUp Then
                ' same_together = True, copyDirection = Down, byRow = True
                '02_01
                    For i = LBound(arr, 1) To UBound(arr, 1)
                        row = A_GetRow(arr, i)
                        For j = 0 To n - 1
                            block = A_Append2(block, row)
                        Next j
                        outArr = A_Extend(outArr, block)
                        Erase block
                    Next i
                Else
                ' same_together = True, copyDirection = Right, byRow = True
                '03_01
                    For i = LBound(arr, 1) To UBound(arr, 1)
                        row = A_GetRow(arr, i)
                        'Dim block() As Variant
                        For j = 0 To n - 1
                            block = A_Append2(block, row)
                        Next j
                        If A_IsEmpty(outArr) Then
                            outArr = block
                        Else
                            outArr = A_HStackH1(outArr, block)
                        End If
                        Erase block
                    Next i
                End If
            Else
                If copyDirection = xlDown Or copyDirection = xlUp Then
                '02_02
                    For i = LBound(arr, 2) To UBound(arr, 2)
                        Vertical = A_GetColumn(arr, i)
                        'Dim block() As Variant
                        For j = 0 To n - 1
                            block = A_HStack(block, Vertical)
                        Next j
                        If A_IsEmpty(outArr) Then
                            outArr = block
                        Else
                            outArr = A_Extend(outArr, block)
                        End If
                        Erase block
                    Next i
                Else
                '03_02
                    For i = LBound(arr, 2) To UBound(arr, 2)
                        Vertical = A_GetColumn(arr, i)
                        'Dim block() As Variant
                        For j = 0 To n - 1
                            block = A_HStack(block, Vertical)
                        Next j
                        If A_IsEmpty(outArr) Then
                            outArr = block
                        Else
                            outArr = A_HStack(outArr, block)
                        End If
                        Erase block
                    Next i
                End If
            End If
        Else
        '**************************************Done
            If copyDirection = xlDown Or copyDirection = xlUp Then
                ' same_together = False, copyDirection = Down
                '01_01
                outArr = A_VStackRep(arr, n)
            Else
                ' same_together = False, copyDirection = Right
                '01_02
                outArr = A_HStackRep(arr, n)
            End If
        End If
        
        
    Else
        MsgBox ("Don't support > 2d array [:From A_Replicates]")
    End If
    A_Replicates = outArr
End Function
Function A_IsEmpty(arr)
'Hard for ChatGPT
    n_dim = A_NDim(arr)
    On Error GoTo Err01
    
    If n_dim = 1 Then
        
        check_str = Trim(Join(arr))
        If check_str = "" Then
            A_IsEmpty = True
        Else
            A_IsEmpty = False
        End If
        
        Exit Function
    ElseIf n_dim = 2 Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If IsEmpty(arr(i, j)) Then
                    A_IsEmpty = True
                    Exit Function
                End If
            Next j
        Next i
        A_IsEmpty = False
        Exit Function
    ElseIf n_dim = 0 Then
        A_IsEmpty = True
        Exit Function
    Else
    
        MsgBox ("Can't check more than 2d array [:From A_IsEmpty]")
    End If
    On Error GoTo 0
    Exit Function
Err01:
    A_IsEmpty = True
End Function


Function Rg_RangeFromText(text_for_row, text_for_col, ws_name, Optional LookAtRow As XlLookAt = xlPart, Optional LookAtCol As XlLookAt = xlPart)
'Assume that in the sheet there's only 1 valide text_for_row & text_for_col
'Assume text_for_row & text_for_col are in the same sheet
'If couldn't find it will return Nothing
    On Error GoTo Err01
    Set row_cell = Rg_FindAllRange(text_for_row, ws_name, LookAtRow)
    Set col_cell = Rg_FindAllRange(text_for_col, ws_name, LookAtCol)
    Set outRng = Rg_RangeFromRange(row_cell, col_cell)
    Set Rg_RangeFromText = outRng
    Exit Function
Err01:
    Set Rg_RangeFromText = Nothing
End Function
Function Rg_RangeFromRange(rng_for_row, rng_for_col)
'Assume rng_for_row& rng_for_col are in the same sheet & are only 1 cell
    Dim row_num As Long
    Dim col_num As Long
    row_num = rng_for_row.row
    col_num = rng_for_col.column
    Set Rg_RangeFromRange = rng_for_row.Parent.Cells(row_num, col_num)
End Function
Function A_ShiftIndexTo0(arr)
    If A_NDim(arr) = 1 Then
        shift = -LBound(arr)
        A_ShiftIndexTo0 = A_ShiftIndex(arr, shift)
    ElseIf A_NDim(arr) = 2 Then
    'Assume that both has index of 1
        shift = -LBound(arr, 1)
        A_ShiftIndexTo0 = A_ShiftIndex(arr, shift)
    Else
        MsgBox ("Don't support >2d array")
    End If
    
End Function
Function A_Extend(arr1, arr2)
    Dim out_arr() As Variant
    If A_IsEmpty(arr1) Then
        A_Extend = arr2
        Exit Function
    End If
    
    If A_IsEmpty(arr2) Then
        A_Extend = arr1
        Exit Function
    End If


    If A_NDim(arr1) = 1 And A_NDim(arr2) = 1 Then
        For Each elem In arr1
            out_arr = A_Append2(out_arr, elem)
        Next
        For Each elem In arr2
            out_arr = A_Append2(out_arr, elem)
        Next
        A_Extend = out_arr
        Exit Function
    ElseIf A_NDim(arr1) = 2 And A_NDim(arr2) = 2 Then
        'out_arr = arr1
        
        arr1 = A_ShiftIndexTo0(arr1)
        arr2 = A_ShiftIndexTo0(arr2)
        numRows1 = UBound(arr1, 1)
        numCols1 = UBound(arr1, 2)
        numRows2 = UBound(arr2, 1)
        numCols2 = UBound(arr2, 2)
        ReDim Preserve out_arr(numRows1 + numRows2 + 1, numCols2)
        
        For i = 0 To numRows1
            For j = 0 To numCols2
                out_arr(i, j) = arr1(i, j)
            Next j
        Next i
        
        For i = 0 To numRows2
            For j = 0 To numCols2
                out_arr(numRows1 + i + 1, j) = arr2(i, j)
            Next j
        Next i
        A_Extend = out_arr
        Exit Function
    Else
        MsgBox ("This function not support >2d array or not Equal dimesion array")
    End If

    
End Function
Function VB_Transpose(arr)

    VB_Transpose = A_ShiftIndexTo0(WorksheetFunction.transpose(arr))
    
End Function
Function A_RepMerge(rng)
    Dim outArr() As Variant
    Dim cell As range
    For Each cell In rng
        'temp_arr = A_RepMerge1Cell(cell)
        If cell.value = "" Then
            val_to_fill = prev_val
        Else
            prev_val = cell.value
            val_to_fill = cell.value
        End If
        outArr = A_Append(outArr, val_to_fill)
        
        
    Next cell
    A_RepMerge = outArr
End Function

Function Rg_Resize(rng, rowsize, ColumnSize) As range
' rowsize As Integer
'ColumnSize As Integer
    Dim curRange As range

    Set curRange = rng.Cells(1, 1)
    Dim newRange As range
    If rowsize >= 0 And ColumnSize >= 0 Then
        Set newRange = curRange.Resize(rowsize, ColumnSize)
    ElseIf rowsize < 0 And ColumnSize < 0 Then
        Set anchor = curRange.Offset(rowsize + 1, ColumnSize + 1)
        Set newRange = anchor.Resize(-rowsize, -ColumnSize)
    ElseIf rowsize < 0 Then
        Set anchor = curRange.Offset(rowsize + 1, 0)
        Set newRange = anchor.Resize(-rowsize, ColumnSize)
    Else
        Set anchor = curRange.Offset(0, ColumnSize + 1)
        Set newRange = anchor.Resize(rowsize, -ColumnSize)
    End If
    
    Set Rg_Resize = newRange
End Function
Function Rg_Resize2(rng, rowsize As Integer, ColumnSize As Integer) As range
    Dim start_cell As range
    Set start_cell = rng.Cells(1, 1)
    
    If rowsize < 0 Then
        Set outRange = start_cell.Resize(, -1 * rowsize).Offset(, ColumnSize)
    ElseIf ColumnSize < 0 Then
        Set outRange = start_cell.Resize(-1 * ColumnSize).Offset(rowsize)
    Else
        Set outRange = start_cell.Resize(rowsize, ColumnSize)
    End If
    Rg_Resize2 = outRange
End Function

Function Rg_NextContainNum(rng, direction As XlDirection) As range
'Return nothing if no cells are found
    ' Declare a variable to store the next cell that has text
    wb_name = rng.Parent.Parent.name
    ws_name = rng.Parent.name
    Dim nextCell As range
    Dim row_offset As Integer
    Dim col_offset As Integer

    ' Set the next cell to the current range
    Set nextCell = Workbooks(wb_name).Sheets(ws_name).range(rng.Address)
    ' Set the offset
    If direction = xlUp Then
        row_offset = -1
        col_offset = 0
    ElseIf direction = xlDown Then
        row_offset = 1
        col_offset = 0
    ElseIf direction = xlToLeft Then
        row_offset = 0
        col_offset = -1
    ElseIf direction = xlToRight Then
        row_offset = 0
        col_offset = 1
    End If
    ' Loop until a cell with text is found or the end of the worksheet is reached
    Set nextCell = nextCell.Offset(row_offset, col_offset)
    On Error GoTo Err01
    Do Until St_ContainsNum(nextCell.value) And nextCell.value <> ""
        ' Move to the next cell in the specified direction
        Set nextCell = nextCell.Offset(row_offset, col_offset)
    Loop
    On Error GoTo 0
    Set Rg_NextContainNum = Workbooks(wb_name).Sheets(ws_name).range(nextCell.Address)
    Exit Function
Err01:
    Set Rg_NextContainNum = Nothing
End Function
Function Rg_NextNumeric(rng, direction As XlDirection) As range
'Return nothing if no cells are found
    ' Declare a variable to store the next cell that has text
    wb_name = rng.Parent.Parent.name
    ws_name = rng.Parent.name
    Dim nextCell As range
    Dim row_offset As Integer
    Dim col_offset As Integer

    ' Set the next cell to the current range
    Set nextCell = Workbooks(wb_name).Sheets(ws_name).range(rng.Address)
    ' Set the offset
    If direction = xlUp Then
        row_offset = -1
        col_offset = 0
    ElseIf direction = xlDown Then
        row_offset = 1
        col_offset = 0
    ElseIf direction = xlToLeft Then
        row_offset = 0
        col_offset = -1
    ElseIf direction = xlToRight Then
        row_offset = 0
        col_offset = 1
    End If
    ' Loop until a cell with text is found or the end of the worksheet is reached
    Set nextCell = nextCell.Offset(row_offset, col_offset)
    On Error GoTo Err01
    Do Until IsNumeric(nextCell) And nextCell.value <> ""
        ' Move to the next cell in the specified direction
        Set nextCell = nextCell.Offset(row_offset, col_offset)
    Loop
    On Error GoTo 0
    Set Rg_NextNumeric = Workbooks(wb_name).Sheets(ws_name).range(nextCell.Address)
    Exit Function
Err01:
    Set Rg_NextNumeric = Nothing
End Function
Function Rg_NextTextCell(rng, direction As XlDirection) As range
'Return nothing if no cells are found
    ' Declare a variable to store the next cell that has text
    wb_name = rng.Parent.Parent.name
    ws_name = rng.Parent.name
    Dim nextTextCell As range
    Dim row_offset As Integer
    Dim col_offset As Integer

    ' Set the next cell to the current range
    Set nextTextCell = Workbooks(wb_name).Sheets(ws_name).range(rng.Address)
    ' Set the offset
    If direction = xlUp Then
        row_offset = -1
        col_offset = 0
    ElseIf direction = xlDown Then
        row_offset = 1
        col_offset = 0
    ElseIf direction = xlToLeft Then
        row_offset = 0
        col_offset = -1
    ElseIf direction = xlToRight Then
        row_offset = 0
        col_offset = 1
    End If
    ' Loop until a cell with text is found or the end of the worksheet is reached
    Set nextTextCell = nextTextCell.Offset(row_offset, col_offset)
    On Error GoTo Err01
    Do Until nextTextCell.value <> ""
        ' Move to the next cell in the specified direction
        Set nextTextCell = nextTextCell.Offset(row_offset, col_offset)
    Loop
    On Error GoTo 0
    Set Rg_NextTextCell = Workbooks(wb_name).Sheets(ws_name).range(nextTextCell.Address)
    Exit Function
Err01:
    Set Rg_NextTextCell = Nothing
End Function
Function Rg_NextNoTextCell(rng, direction As XlDirection) As range
    ' Declare a variable to store the next blank cell
    
    Dim nextBlankCell As range
    Dim row_offset As Integer
    Dim col_offset As Integer
    wb_name = rng.Parent.Parent.name
    ws_name = rng.Parent.name
    ' Set the next blank cell to the current range
    Set nextBlankCell = Workbooks(wb_name).Sheets(ws_name).range(rng.Address)
    ' Set the offset
    If direction = xlUp Then
        row_offset = -1
        col_offset = 0
    ElseIf direction = xlDown Then
        row_offset = 1
        col_offset = 0
    ElseIf direction = xlToLeft Then
        row_offset = 0
        col_offset = -1
    ElseIf direction = xlToRight Then
        row_offset = 0
        col_offset = 1
    End If
    ' Loop until a blank cell is found or the end of the worksheet is reached
    Do Until IsEmpty(nextBlankCell) Or nextBlankCell Is Nothing
        ' Move to the next cell in the specified direction
        Set nextBlankCell = nextBlankCell.Offset(row_offset, col_offset)
    Loop
    
    ' Return the next blank cell
    Set Rg_NextNoTextCell = Workbooks(wb_name).Sheets(ws_name).range(nextBlankCell.Address)

End Function

Function Rg_ReorderRange(rng)
'Reorder range by their addresses when it's not in order
    ws_name = rng.Parent.name
    Set wb = rng.Parent.Parent
    Dim ws As Worksheet
    Set ws = wb.Worksheets(ws_name)
    
    Dim out_rng As range
    Dim addresses() As String
    ReDim addresses(0 To rng.Cells.count - 1) As String
    Dim i As Long
    i = 0
    For Each cell In rng
        addresses(i) = cell.Address
        i = i + 1
    Next cell
    ' sort the addresses array
    Dim temp As String
    For i = 0 To UBound(addresses) - 1
        For j = i + 1 To UBound(addresses)
            If addresses(i) > addresses(j) Then
                temp = addresses(i)
                addresses(i) = addresses(j)
                addresses(j) = temp
            End If
        Next j
    Next i
    Set out_rng = ws.range(addresses(0))
    For i = 1 To UBound(addresses)
        Set out_rng = Union(out_rng, ws.range(addresses(i)))
    Next i
    Set Rg_ReorderRange = out_rng
End Function

Function Rg_FindSomeRanges(searchArr As Variant, inx_arr, Optional ws = "", Optional each_cell = True, Optional myLookAt As XlLookAt = xlPart, Optional wb = "")
    'searchString: array or string
'Return the ranges with only some index using 1-index
'[2] => 2nd element
'2 => 2 ranges
'-3 => 3 last ranges
'[-1,-2,-4] => last, 2nd last and 4th last
  '**********************************************************Paramenters******************'
'each_cell = True  => Loop through "individual cell"
'each_cell = False  => Loop through each "area"
    '**********************************************************Handle worksheet both type ws and string******************'
    Set targetSheet = Ws_WS_at_WB(ws, wb)
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------'
    Set AllRange = Rg_FindAllRange(searchArr, ws, myLookAt, wb)
    n_area = AllRange.Areas.count
    n_cell = AllRange.count
    count = 1
    Dim outRange As range
    If each_cell Then
        '**********************************************************Case Loop individual cell******************'
        If Not IsArray(inx_arr) Then
              '**********************************************************Case inx_arr is number******************'
            For Each curr_cell In AllRange
                from_inx = n_cell + inx_arr + 1
                If inx_arr > 0 Then
                    If count <= inx_arr Then
                        Set outRange = Rg_Union(outRange, curr_cell)
                    End If
                Else
                    If from_inx <= count Then
                        Set outRange = Rg_Union(outRange, curr_cell)
                    End If
                End If
                count = count + 1
            Next curr_cell
            '**********************************************************Case inx_arr is array*****************'
            
              
        
        Else
            For Each curr_cell In AllRange
                For i = LBound(inx_arr) To UBound(inx_arr)
                    from_inx = n_cell + inx_arr(i) + 1
                    If inx_arr(i) > 0 Then
                        If count = inx_arr(i) Then
                            Set outRange = Rg_Union(outRange, curr_cell)
                        End If
                    Else
                        If from_inx = count Then
                            Set outRange = Rg_Union(outRange, curr_cell)
                        End If
                    End If
                Next i
                
                count = count + 1
            Next curr_cell
            
            ' For Each curr_area In AllRange.Areas
            '     from_inx = n_area + inx_arr
            '     If inx_arr > 0 Then
            '         If count <= inx_arr Then
            '             Set outRange = Rg_Union(outRange, curr_area)
            '         End If
            '     Else
            '         If from_inx <= count Then
            '             Set outRange = Rg_Union(outRange, curr_area)
            '         End If
            '     End If
            '     count = count + 1
            ' Next curr_area
        End If
            
        '**********************************************************each_cell = True*************************************'
    Else
     '**********************************************************Case Loop through area******************'
        '**********************************************************each_cell = False*************************************'
        If IsArray(inx_arr) Then
            '**********************************************************Case inx_arr is array******************'
            For i = LBound(inx_arr) To UBound(inx_arr)
                If inx_arr(i) > 0 Then
                    curr_area = AllRange.Areas(inx_arr(i))
                Else
                    back_inx = n_area + inx_arr(i) + 1
                    curr_area = AllRange.Areas(back_inx)
                End If
            Next i
            Set outRange = Rg_Union(outRange, curr_area)
        Else
         '**********************************************************Case inx_arr is numbers******************'
             For Each curr_area In AllRange.Areas
                from_inx = n_area + inx_arr + 1
                If inx_arr > 0 Then
                    If count <= inx_arr Then
                        Set outRange = Rg_Union(outRange, curr_area)
                    End If
                    If from_inx <= count Then
                        Set outRange = Rg_Union(outRange, curr_area)
                    End If
                End If
                count = count + 1
            Next curr_area
         
        End If
        
        
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------'
    End If
    Set outRange2 = Rg_ReorderRange(outRange)
    Set Rg_FindSomeRanges = outRange2
End Function
Function Rg_FindAllRange(searchArr As Variant, Optional ws = "", Optional myLookAt As XlLookAt = xlPart, Optional wb = "") As range
'Hard for ChatGPT
    
    Set targetSheet = Ws_WS_at_WB(ws, wb)

    Dim outRange As range
    
    If IsArray(searchArr) Then
        For i = LBound(searchArr) To UBound(searchArr)
            Set currFound = Rg_FindAllRangeH1(searchArr(i), ws, myLookAt, wb)
            Set outRange = Rg_Union(outRange, currFound)
        Next
    Else
        Set outRange = Rg_FindAllRangeH1(searchArr, ws, myLookAt, wb)
    End If
    
    Set Rg_FindAllRange = outRange
End Function

Function Rg_FindAllRangeH1(searchString, Optional ws = "", Optional myLookAt As XlLookAt = xlPart, Optional wb = "")
'It works
'Hard for ChatGPT
    Set targetSheet = Ws_WS_at_WB(ws, wb)
    Dim SearchRange As range
    Set SearchRange = targetSheet.usedRange ' adjust the range to your needs

    Dim foundRange As range
    Set foundRange = SearchRange.Find(searchString, LookIn:=xlValues, LookAt:=myLookAt)
    
    Set resFind0 = SearchRange.Find(What:=searchString, LookIn:=xlValues, LookAt:=myLookAt)
    Set curr_found = resFind0
    Dim allFoundRanges As range
    
    If resFind0 Is Nothing Then
        Set Rg_FindAllRangeH1 = Nothing
        Exit Function
    Else
        Set allFoundRanges = resFind0
    End If
    first_address = resFind0.Address
    
    
    Do While True
        Set curr_found = SearchRange.Find(What:=searchString, after:=curr_found, LookIn:=xlValues, LookAt:=myLookAt)
        found_sheet_name = curr_found.Parent.name
        found_address = curr_found.Address
        saved_address = found_sheet_name & "   " & found_address
        
        
        If curr_found.Address = first_address Then
            Set Rg_FindAllRangeH1 = allFoundRanges
            Exit Do
        Else
            Set allFoundRanges = Union(allFoundRanges, curr_found)
        End If
       
    Loop
    
    
End Function
Function St_FileNameFromPath(filepath) As String
    Dim fileName As String
    fileName = Split(filepath, "\")(UBound(Split(filepath, "\")))
    St_FileNameFromPath = fileName
End Function

Function W_SMCode(prodCode)
    Dim ws01 As Worksheet
    Dim prodCodeList, SMCode As range
    Set ws01 = ThisWorkbook.Sheets("Master")
    Set prodCodeList = ws01.range("B:B")
    Set SMCode = ws01.range("D:D")
    If IsNumeric(prodCode) Then
        prodCode = CLng(prodCode)
    End If
    res = VB_Xlookup(prodCode, prodCodeList, SMCode)
    W_SMCode = res
    
    
End Function
Sub Wb_ColorSheetTabsAfter(wb, inx, Optional tab_color = -0.5)
'Create the similar Sub with Color After some Name
'Before some index
'Mid index(between)
'Right
'Or between 2 sheets names(in string)
' Or delete tab color
    If tab_color = -0.5 Then
        tab_color = RGB(255, 153, 0)
    End If
    For i = 1 To wb.Sheets.count
        If i >= inx Then
            wb.Sheets(i).Tab.color = tab_color 'Orange color
        End If
    Next i
End Sub

Sub Wb_CheckNumSheets(wb, threshold As Integer)
    If wb.Sheets.count >= threshold Then
        MsgBox Workbook.name & " has " & wb.Sheets.count & " sheets."
    End If
End Sub

Function OS_OpenExcelLatest(folderPath)
 '--- change this to sub called OS_GetLatestExcelFile
    mostRecentFile = ""
    mostRecentFileDate = #1/1/1900#
    Dim fso, folder, file As Object
    Dim wb_result As Workbook
    ' Loop through the files in the folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If LCase(Right(file.name, 4)) = ".xls" Or LCase(Right(file.name, 5)) = ".xlsx" Or LCase(Right(file.name, 5)) = ".xlsb" Or LCase(Right(file.name, 5)) = ".xlsm" Then
            If file.DateLastModified > mostRecentFileDate Then
                mostRecentFile = file.path
                mostRecentFileDate = file.DateLastModified
            End If
        End If
    Next file
    
    ' Open the most recent file
    On Error GoTo Err01
    'Don't update links when open
    Set wb_result = Workbooks.Open(fileName:=mostRecentFile, UpdateLinks:=0)
    On Error GoTo 0
    Set OS_OpenExcelLatest = wb_result
    Exit Function
Err01:
    'myFileName = St_FileNameFromPath(folderPath)
    'MsgBox (myFileName & "Has no Excel File")
    MsgBox ("Has no Excel File")
    'OS_OpenExcelLatest = "NoExcelFile"
    Resume Next
    Exit Function
    
End Function
Function St_GetBefore(str, chr_list As Variant)
    outStr = str
    If IsArray(chr_list) Then
        For i = LBound(chr_list) To UBound(chr_list)
            outStr = St_GetBeforeH1(outStr, chr_list(i))
        Next
        St_GetBefore = outStr
    Else
        St_GetBefore = St_GetBeforeH1(str, chr_list)
    End If
End Function

Function St_GetBeforeH1(str, chr) As String
    Dim charPos As Long
    charPos = InStr(str, chr)
    If charPos > 0 Then
        St_GetBeforeH1 = Left(str, charPos - 1)
    Else
        St_GetBeforeH1 = str
    End If
End Function


Sub W_GetTemplateExample()
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' 1st Finish at 2 hr 15 min
'Not Done Yet 3 hr + Of Debugging
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileName As String
'folder path of wb_numbers (file that P.Oat fill it manually)
    folderPath = "C:\Users\n1603499\OneDrive - Liberty Mutual\Documents\12.02  Rotation2  EastProduct\Test01"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Dim out_arr() As Variant
    Dim wb_result, wb_campaign, wb_numbers As Workbook
    Dim n_sheetLimit As Integer
'Set high numbers if you wanna run all files
'This limit the # of files opened to debug
    n_file_limit = 1000
    n_sheetLimit = 6
    
    output_folder_path = "C:\Users\n1603499\OneDrive - Liberty Mutual\Documents\12.02  Rotation2  EastProduct\Test01\OutPutFile"
    'Watching index when debugging
    count = 0
    For Each file In folder.Files
    'Watching index when debugging
        count = count + 1
        If count > n_file_limit Then
            Exit For
        End If
        If LCase(Right(file.name, 4)) = ".xls" Or LCase(Right(file.name, 5)) = ".xlsx" Or LCase(Right(file.name, 5)) = ".xlsb" Or LCase(Right(file.name, 5)) = ".xlsm" Then
            delimiter = Array("_", " ")
            prodCode = St_GetBefore(Left(file.name, (Len(file.name) - 5)), delimiter)
            
            SMCode = W_SMCode(prodCode)
            outputName = prodCode & "_" & SMCode
            output_path = output_folder_path & "\" & outputName
            Set wb_result = Workbooks.Add
            'Set wb_result = Workbooks.Open("C:\Users\n1603499\OneDrive - Liberty Mutual\Desktop\VBA LibFile V06.02.xlsb")
 


'#####################Template Part############################

            Set wb_numbers = Workbooks.Open(file.path)
            
            Set sheet01 = wb_numbers.Sheets("Coverage Input")
            Set sheet02 = wb_numbers.Sheets("Net Premium Input")
            Set sheet03 = wb_numbers.Sheets("Eligibility Input")
            
            Application.DisplayAlerts = False
            ' Copy the sheets to the wb_result workbook
            sheet01.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            sheet02.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            sheet03.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            Application.DisplayAlerts = True
            
            ' Change the tab color to green
            For Each ws_temp In wb_result.Sheets
                ws_temp.Tab.color = RGB(0, 255, 0)
            Next
            'wb_result.Sheets(wb_result.Sheets.count).Tab.color = RGB(0, 255, 0)
            'wb_result.Sheets(wb_result.Sheets.count - 1).Tab.color = RGB(0, 255, 0)
            'wb_result.Sheets(wb_result.Sheets.count - 2).Tab.color = RGB(0, 255, 0)
            
'----------------------------------------Template Part-----------------------------------------
'#####################campaign Part############################
            campaign_folder_path = W_FindClaimPath2(prodCode)
            
            Set page01 = wb_result.Worksheets("Sheet1")
            page01.name = "FolderPath"
            page01.range("A1") = campaign_folder_path
            
            On Error GoTo Err01
            Set wb_campaign = OS_OpenExcelLatest(campaign_folder_path)
            
            If wb_campaign.Sheets.count > n_sheetLimit Then
' Check this line
'Add the name to indicate that this product code has too many sheets in campaign files
                output_path = output_path & "_ManySheets"
                Application.DisplayAlerts = False
                wb_result.SaveAs (output_path)
                Application.DisplayAlerts = True
                Call Wb_CheckNumSheets(wb_campaign, n_sheetLimit)
                GoTo SkipIteration
            End If
           'Save as without asking
            Application.DisplayAlerts = False
            wb_result.SaveAs (output_path)
            Application.DisplayAlerts = True
            
            
            oldSheetNum = wb_result.Sheets.count
            ColorOrange = CLng("&H" & "FF6600")
            On Error GoTo 0
            
            wb_result.Activate
            Application.DisplayAlerts = False
            For Each ws In wb_campaign.Sheets
                ws.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            Next ws
            Call Wb_ColorSheetTabsAfter(wb_result, oldSheetNum + 1, ColorOrange)

            'ThisWorkbook.Activate
SkipIteration:
            wb_numbers.Close SaveChanges:=False
            wb_campaign.Close SaveChanges:=False
            Application.DisplayAlerts = True
            
            wb_result.Close
Back01:
'----------------------------------------campaign Part-----------------------------------------
            ' ############################################################## Get Lastest Campaign Excel File ##################################################################
           
             ' ----------------------------------------------------------------------------- Get Most Recent Campaign Excel File ---------------------------------------------------------------------------------------------
            
            
        End If
    Next file
    Exit Sub
Err01:
    'newFileName = wb_result.name & "_NoExcelCampaignFile"
    output_path = output_path & "_NoExcelCampaignFile"
    Application.DisplayAlerts = False
    wb_result.SaveAs (output_path)
    Application.DisplayAlerts = True
            
    wb_numbers.Close
    wb_result.Close
    Resume Back01
    
End Sub
