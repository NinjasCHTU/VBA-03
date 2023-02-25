Attribute VB_Name = "Func_SpaceANDSelect1"
Function Sp_GroupByColor()
    Application.Caller.Interior.color = vbYellow
End Function
Function Sp_PickColumn(table, col_n)
    Sp_PickColumn = table.Columns(col_n)
    

End Function

Function Sp_PickRow(table, row_n)
    Sp_PickRow = table.Rows(row_n)
    
End Function

Function Sp_Vstack(selectRange)
'@@@@@@@@@@@@@@@@Dependency -> Array1
    Dim outArr() As Variant
    n_row = selectRange.Rows.Count
    n_col = selectRange.Columns.Count
    

    
    For i = 1 To n_col
        For Each curr_temp In selectRange.Columns(i)
            curr_col = curr_temp.value
            
            
            For j = 1 To n_row
                curr_elem = curr_col(j, 1)
                outArr = A_Append(outArr, curr_elem)
            Next j
            
            
            
        Next curr_temp
    Next i
    
    
    
    Sp_Vstack = VB_Transpose(outArr)
    
End Function
Function Sp_PickFromNColor(arr_in As range, ParamArray color_list())

End Function
Function Sp_PickFrom1Color(arr_in As range, color As range, Optional direction = 0)
    
'@@@@@@@@@@@@@@@@Dependency -> A_Append, Lib_Array1
    myColor = color.Interior.colorIndex
    Dim out_arr() As Variant
    
    
    For Each curr_cell In arr_in
        curr_val = curr_cell.value
        curr_color = curr_cell.Interior.colorIndex
        If curr_color = myColor Then
            out_arr = A_Append(out_arr, curr_val)
        End If
    Next
    If direction = 0 Then
    
        Sp_PickFrom1Color = VB_Transpose(out_arr)
    Else
        Sp_PickFrom1Color = out_arr
    End If
    ActiveSheet.Calculate
    

End Function
Function Sp_PickFromAnyColor(arr_in As range)
   
' Color_index of NOTFILL is -4142

'@@@@@@@@@@@@@@@@Dependency -> A_Append, Lib_Array1
    
    Dim out_arr() As Variant
    
    
    For Each curr_cell In arr_in
        curr_val = curr_cell.value
        curr_color = curr_cell.Interior.colorIndex
        If curr_color <> -4142 Then
            out_arr = A_Append(out_arr, curr_val)
        End If
    Next
    If direction = 0 Then
        Sp_PickFromAnyColor = VB_Transpose(out_arr)
    Else
        Sp_PickFromAnyColor = out_arr
    End If
    Application.Volatile

End Function
Function Sp_SelectFromTL(inCell, Optional n_row = 1, Optional n_col = 1)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim outRange As range
    
    upLeft_row = inCell.row
    upLeft_col = inCell.column
    
    'temp = Range(Cells(inCell.row, inCell.col), Cells(inCell.row + n_row, inCell.column + n_col))
    
    Set outRange = range(Cells(inCell.row, inCell.column), Cells(inCell.row + n_row - 1, inCell.column + n_col - 1))
    'Set outRange = Range(Cells(6, 1), Cells(8, 2))
    'Sp_SelectFromTL = inCell.row
    Set Sp_SelectFromTL = outRange
    
End Function

Function Sp_CombineToV(ParamArray arr_range() As Variant)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim temp_arr() As Variant
    Dim i As Integer
    i = 0
    For Each range_col In arr_range
        For Each elem In range_col
            ReDim Preserve temp_arr(0 To i)
            temp_arr(i) = elem
            i = i + 1
        Next
    Next
    Sp_CombineToV = WorksheetFunction.Transpose(temp_arr)
End Function

Function Sp_CombineToH(ParamArray arr_range() As Variant)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim temp_arr() As Variant
    Dim i As Integer
    i = 0
    For Each range_col In arr_range
        For Each elem In range_col
            ReDim Preserve temp_arr(0 To i)
            temp_arr(i) = elem
            i = i + 1
        Next
    Next
    Sp_CombineToH = (temp_arr)
End Function
Function Sp_to1DLine(inRange As range, Optional direction = 0)
    n_area = inRange.Count
    Dim out_arr() As Variant
    ReDim Preserve out_arr(n_area - 1)
    
    i = 0
    If direction = 0 Then
        For Each curr_col In inRange.Columns
            For Each elem In curr_col.Value2
                out_arr(i) = elem
                i = i + 1
            Next
        Next
    Else
        For Each curr_row In inRange.Rows
            For Each elem In curr_row.Value2
                out_arr(i) = elem
                i = i + 1
            Next
        Next
    End If
    

    
    Sp_to1DLine = VB_Transpose(out_arr)
    
    
    

End Function
Function Sp_to2DTable(row, col, Optional direction = 1)
    '@@@@@@@@@@@@@@@@Dependency ->
End Function
Function Sp_to3DLayer()

End Function

'This is not Done
Function Sp_toDiagonal(arr_in)
    '@@@@@@@@@@@@@@@@Dependency ->
    n = arr_in.Count
    Dim arr() As Variant
    ReDim arr(n, n)
    i = 0
    For Each elem In arr_in
        arr(i, i) = arr_in(i)
        i = i + 1
    Next
    
    Sp_toDiagonal = arr
    

End Function

Function Sp_SelectSkipVB(arr, r, n)
Attribute Sp_SelectSkipVB.VB_Description = "The generalized select skip.    If you want to select 1 line skip 1 line then r should be 0 or 1  and n =2"
Attribute Sp_SelectSkipVB.VB_ProcData.VB_Invoke_Func = " \n14"
    '@@@@@@@@@@@@@@@@Dependency ->
    On Error GoTo ErrCase
    n_arr = arr.Count
    m = n_arr \ n
    If r = n Then
        r = 0
    End If
    Dim temp_arr() As Variant
    ReDim temp_arr(m - 1)
    arr_i = 0
    
    For i = 1 To n_arr
        curr_r = i Mod n
        If curr_r = r Then
            
            temp_arr(arr_i) = arr(i)
            arr_i = arr_i + 1
        End If
    Next i

Done:
    Sp_SelectSkipVB = VB_Transpose(temp_arr)
    Exit Function
ErrCase:
    Sp_SelectSkipVB = "n must divide the #rows of this array"
    'r = is the remander selected r must be less than n
    'n = is the # of repeated cycles
    
End Function



