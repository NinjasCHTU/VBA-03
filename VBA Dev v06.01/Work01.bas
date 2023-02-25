Attribute VB_Name = "Work01"
Sub W_MakeCookieReport()
    Call W_CopySplitColumn("VT", "VT2", 5, "C", "Delegated Attribute", "Applicable to all levels and products")
    Call W_CopySplitColumn("HK", "HK2", 6, "D", "Delegated Attribute", "Applicable to all levels and products")
    Call W_CopySplitColumn("SG", "SG2", 7, "C", "Delegated Attribute", "Applicable to all levels and products")
    Call W_CopySplitColumn("TH", "TH2", 16, "C", "Delegated Attribute", "Applicable to all levels and products")
    MsgBox ("DONE !!!!")
    
End Sub

Sub W_CopySplitColumn(input_shName, output_shName, n_col, start_col, split_by, end_with)
    Dim currSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim outputRange As range
    
    'split_by = split_by.value
    'end_with = end_with.value
    Set currSheet = Worksheets(input_shName)
    row_list = Rg_SearchRowNum(input_shName, split_by)
    end_row = Rg_SearchRowNum(input_shName, end_with)
    end_row(0) = end_row(0) - 1
    row_list = W_GetAreaInx(row_list)
    row_list = A_Extend(row_list, end_row)
    
    Dim FirstCol As range
    
    n_area = CInt(UBound(row_list) / 2) + 1
    'Set FirstCol = Worksheets("TH").Range("C10:C21")
    'Set FirstCol = Union(FirstCol, Worksheets("TH").Range("C24:C31"))
    'Set FirstCol = Rg_GetRangeSameColManyRows("TH", "C", 10, 21, 24, 31, 34, 49)
    Set FirstCol = Rg_GetRangeSameColManyRows(input_shName, start_col, row_list)
    'Set FirstCol = Union(FirstCol, Worksheets("TH").Range("C10:C21"))
    
    'Set FirstCol = currSheet.Range(Range("C10"), Range("C21"))
    Set outputSheet = Worksheets(output_shName)
    Set outputFirst = outputSheet.range("A1")
    
    'FirstCol.Copy outputFirst
    'outputSheet.Columns("A").AutoFit
    
    
    Dim range_arr As Variant
    Dim curr_range As range
    
    Dim output_big_range As range
    
    For i = 1 To n_col
        If output_big_range Is Nothing Then
            Set output_big_range = FirstCol.Offset(0, i)
        Else
            Set output_big_range = Union(output_big_range, FirstCol.Offset(0, i))
        End If
        Set curr_range = FirstCol.Offset(0, i)
        range_arr = A_Append(range_arr, curr_range)
    Next
    
    i = 0
    n_arr = UBound(range_arr)
    'For i = 0 To n_arr
        'curr_range = range_arr(i)
        'curr_output = outputFirst.Offset(0, 6 * i)
        'FirstCol.Copy curr_output
        'curr_range.Copy curr_output.Offset(0, 1)
    'Next
    
    n_space = 4
    n_area = n_area - 1
    Dim output2 As range
    For i = 0 To n_col
        n_total = 1
        For j = 1 To n_area
            Set curr_range = output_big_range.Areas(j).Columns(i + 1)
            Set curr_output = outputFirst.Offset(0, n_space * i)
'https://www.exceldemy.com/excel-vba-paste-special-values-and-formats/#:~:text=Apply%20InputBox%20in%20VBA%20Paste%20Special%20to%20Copy,Step%205%3A%20Another%20dialog%20box%20will%20appear.%20
'Pastes only value and format
            FirstCol.Copy
            curr_output.PasteSpecial xlPasteValuesAndNumberFormats
            curr_output.PasteSpecial xlPasteFormats
            
            Set output2 = outputSheet.Cells(1, 1)
            Set output2 = outputSheet.Cells(n_total, n_space * i + 2)
            n_section = curr_range.Rows.Count
            n_total = n_total + n_section
            
            'If output2 Is Nothing Then
                'Set output2 = Rg_GetLastUsedRow2("Sheet1", 6 * i + 2)
            'Else
                'Set output2 = Rg_GetLastUsedRow2("Sheet1", 6 * i + 2).Offset(1, 0)
            'End If
            curr_range.Copy
            output2.PasteSpecial xlPasteValuesAndNumberFormats
            output2.PasteSpecial xlPasteFormats
            
            
            curr_output.Columns.ColumnWidth = 50
            outputSheet.Columns(n_space * i + 2).ColumnWidth = 21
        Next
    Next

End Sub
Function W_GetAreaInx(arr_in)
    n = UBound(arr_in)
    Dim out_arr() As Variant
    out_arr = A_Append(out_arr, arr_in(0))
    For i = 1 To n
        curr_val = arr_in(i)
        bottom_val = curr_val - 2
        out_arr = A_Append(out_arr, bottom_val)
        out_arr = A_Append(out_arr, curr_val)
    Next
    W_GetAreaInx = out_arr
End Function
