Attribute VB_Name = "Func_Range01"
Function Rg_ChangeFillColors(rng, color_arr1, color_arr2)
    Dim cell As range
    Dim i As Long
    
    For Each cell In rng
        For i = LBound(color_arr1) To UBound(color_arr1)
            If cell.Interior.color = color_arr1(i) Then
                cell.Interior.color = color_arr2(i)
                Exit For
            End If
        Next i
    Next cell
End Function
Function Rg_ChangeFontColors(rng, color_arr1, color_arr2)
    Dim cell As range
    Dim i As Long
    
    For Each cell In rng
        For i = LBound(color_arr1) To UBound(color_arr1)
            If cell.Font.color = color_arr1(i) Then
                cell.Font.color = color_arr2(i)
                Exit For
            End If
        Next i
    Next cell
End Function

Function Rg_ExtractFontColor(range)
    Dim cell As range
    Dim colorCodes_arr() As Variant
    
    For Each cell In range
        If cell.Font.color <> xlNone And cell.Font.color <> 16777215 Then
            curr_fill_num = cell.Font.color
            
            If Not A_isInArr(colorCodes_arr, curr_fill_num) Then
                colorCodes_arr = A_Append(colorCodes_arr, curr_fill_num)
            End If
        End If
    Next cell
    
    Rg_ExtractFontColor = colorCodes_arr
End Function


Function Rg_ExtractFillColor(range)
    Dim cell As range
    Dim colorCodes_arr() As Variant
    
    For Each cell In range
        If cell.Interior.color <> xlNone And cell.Interior.color <> 16777215 Then
            curr_fill_num = cell.Interior.color
            
            If Not A_isInArr(colorCodes_arr, curr_fill_num) Then
                colorCodes_arr = A_Append(colorCodes_arr, curr_fill_num)
            End If
        End If
    Next cell
    
    Rg_ExtractFillColor = colorCodes_arr
End Function


Function Rg_IfTop(cell As range, value_if_true As String, value_if_false As String)
    If IsEmpty(cell.Offset(-1, 0)) And Not IsEmpty(cell.Offset(1, 0)) Then
        Rg_IfTop = value_if_true
    Else
        Rg_IfTop = value_if_false
    End If
End Function

Function Rg_IsTop(cell As range) As Boolean
    If IsEmpty(cell.Offset(-1, 0)) And Not IsEmpty(cell.Offset(1, 0)) Then
        Rg_IsTop = True
    Else
        Rg_IsTop = False
    End If
End Function
Function Rg_IsGrey(rng) As Boolean
  'Declare variables
  Dim cell As range
  
  'Set Rg_IsGrey to False by default
  Rg_IsGrey = False
  
  'Loop through each cell in the range
  For Each cell In rng
    'Check if the cell's fill color is any shade of grey
    backGroundColor = cell.Interior.color
    
    red = backGroundColor Mod 256
    green = backGroundColor \ 256 Mod 256
    blue = backGroundColor \ 65536 Mod 256
    
    red = CLng(red)
    green = CLng(green)
    blue = CLng(blue)
    
    
    If red = green And green = blue Then
      'If the 3 numbers in the RGB color value are equal, set Rg_IsGrey to True and exit the loop
      Rg_IsGrey = True
      Exit For
    End If
  Next
End Function


Function Rg_IfBottom(cell As range, value_if_true As String, value_if_false As String)
    If Not IsEmpty(cell.Offset(-1, 0)) And IsEmpty(cell.Offset(1, 0)) Then
        Rg_IfBottom = value_if_true
    Else
        Rg_IfBottom = value_if_false
    End If
End Function

Function Rg_IsBottom(cell As range) As Boolean
    If Not IsEmpty(cell.Offset(-1, 0)) And IsEmpty(cell.Offset(1, 0)) Then
        Rg_IsBottom = True
    Else
        Rg_IsBottom = False
    End If
End Function

Function Rg_HasColorCells(range_in)
'ChatGPT
    For Each cell In range_in
        If cell.Interior.colorIndex <> -4142 And cell.Interior.colorIndex <> 7 Then
            Rg_HasColorCells = True
            Exit Function
        End If
    Next cell
    
    Rg_HasColorCells = False
End Function

Function Rg_HasUnfilledCells(range_in)
'ChatGPT
    For Each cell In range_in
        If cell.Interior.colorIndex = -4142 Or cell.Interior.colorIndex = 7 Then
            Rg_HasUnfilledCells = True
            Exit Function
        End If
    Next cell
    
    Rg_HasUnfilledCells = False
End Function
Function Rg_GetRangeSameColManyRows(SheetName, col_name, ParamArray row_inx() As Variant)
    Dim row_inx2() As Variant
    If UBound(row_inx) Mod 2 = 0 Then
        If UBound(row_inx) = 0 Then
            
            n = 0
            For Each temp_arr In row_inx
                For Each MyVal In temp_arr
                    row_inx2 = A_Append(row_inx2, MyVal)
                    
                Next
                
            Next
            
        Else
            MsgBox ("Please Enter the correct number of columns")
        End If
    Else
        row_inx2 = row_inx
    End If
    n = UBound(row_inx2)
    
    m = WorksheetFunction.RoundDown(n / 2, 0)
    Dim out_range As range
    Dim curr_range As range
    For i = 0 To m
        curr_inx = 2 * i
        row_str1 = col_name & row_inx2(curr_inx)
        row_str2 = col_name & row_inx2(curr_inx + 1)
        Set curr_range = Worksheets(SheetName).range(row_str1, row_str2)
        If out_range Is Nothing Then
            Set out_range = curr_range
        Else
            Set out_range = Union(out_range, curr_range)
        End If
    Next
    Set Rg_GetRangeSameColManyRows = out_range
    
End Function
Function Rg_SearchRowNum(shName, search_str, Optional match_case = 0)
'match_case = 0 :> lower or upper letters doesn't matter
'match_case = 1 :> Exact Match
    Dim out_arr() As Variant
    Dim myWorkSh As Worksheet
    Dim myUseRg As range
    Set myWorkSh = Worksheets(shName)
    Set myUseRg = myWorkSh.usedRange
On Error GoTo Err01:

    For Each curr_cell In myUseRg
        If IsError(curr_cell.value) Then
            curr_val = "NA"
        Else
            curr_val = curr_cell.value
        End If
        

        
        If match_case = 1 Then
            curr_val = LCase(curr_val)
            search_str = LCase(search_str)
        End If
        
        If curr_val = search_str Then
            row_inx = curr_cell.row
            out_arr = A_Append(out_arr, row_inx)
        End If
    Next
    
    Rg_SearchRowNum = out_arr

Err01:
'#N/A value
    'MsgBox ("I got this error")
    'MsgBox (curr_cell.Address)
    curr_val = "NA"
    
    

End Function
Function Rg_GetLastUsedRow(column_in)
'This Function return row1 if no data exist in that column
    'Add: Column can be number or Alphabet
    Last_Row = Cells(Rows.count, column_in).End(xlUp).row
    Rg_GetLastUsedRow = Last_Row
    Dim out_range As range
    Set out_range = Cells(Last_Row, column_in)
    Set Rg_GetLastUsedRow = out_range
    
    
End Function

Function Rg_GetLastUsedRow2(SheetName, column_in)
'This Function return row1 if no data exist in that column
    'Add: Column can be number or Alphabet
    Last_Row = Sheets(SheetName).Cells(Rows.count, column_in).End(xlUp).row
    'Rg_GetLastUsedRow2 = Last_Row
    Dim out_range As range
    Set out_range = Sheets(SheetName).Cells(Last_Row, column_in)
    Set Rg_GetLastUsedRow2 = out_range
    
    
End Function
