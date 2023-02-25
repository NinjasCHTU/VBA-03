Attribute VB_Name = "Func_Dear2"
Function D_DictCount(rng As range) As Object
    Dim cell As range
    Dim elementCount As Object
    Set elementCount = CreateObject("Scripting.Dictionary")
    For Each cell In rng
        If Not elementCount.Exists(cell.value) Then
            elementCount.Add cell.value, 1
        Else
            elementCount(cell.value) = elementCount(cell.value) + 1
        End If
    Next cell
    Set D_DictCount = elementCount
End Function


Function D_ToPython1D(rng As range, Optional direction As String = "down") As String
'From ChatGPT
    Dim result As String
    Dim i As Long, j As Long
    
    result = "["
    
    If direction = "down" Then
        For i = 1 To rng.Rows.Count
            For j = 1 To rng.Columns.Count
                If IsNumeric(rng.Cells(i, j).value) Then
                    result = result & rng.Cells(i, j).value & ", "
                Else
                    result = result & """" & rng.Cells(i, j).value & """, "
                End If
            Next j
        Next i
    ElseIf direction = "right" Then
        For j = 1 To rng.Columns.Count
            For i = 1 To rng.Rows.Count
                If IsNumeric(rng.Cells(j, i).value) Then
                    result = result & rng.Cells(j, i).value & ", "
                Else
                    result = result & """" & rng.Cells(j, i).value & """, "
                End If
            Next i
        Next j
    Else
        D_ToPython1D = "The direction should be either ""down"" or ""right"" "
        Exit Function
    End If
    
    result = Left(result, Len(result) - 2) & "]"
    
    D_ToPython1D = result
End Function
Function D_ToPython2D(rng As range) As String
' From ChatGPT
    Dim result As String
    Dim i As Long, j As Long
    
    result = "["
    
    For i = 1 To rng.Rows.Count
        result = result & "["
        For j = 1 To rng.Columns.Count
            If IsNumeric(rng.Cells(i, j).value) Then
                result = result & rng.Cells(i, j).value & ", "
            Else
                result = result & """" & rng.Cells(i, j).value & """, "
            End If
        Next j
        result = Left(result, Len(result) - 2) & "], "
    Next i
    
    result = Left(result, Len(result) - 2) & "]"
    
    D_ToPython2D = result
End Function

Function D_XLookupMany(value_arr, lookup_arr, return_arr)
'Use VStack in value_arr
    lookup_arr = A_toArray2d(lookup_arr)
    return_arr = A_toArray1d(return_arr)
    value_arr = A_Reshape_2dTo1D(value_arr)
    n = UBound(value_arr)
    n_row = UBound(lookup_arr)
    For i = 0 To n_row
        check = 0
        For j = 0 To n
            curr_elem = value_arr(j)
            If lookup_arr(i, j) <> curr_elem Then
                check = 0
                Exit For
            Else
                check = 1
            End If
        Next
        If check = 1 Then
            return_elem = return_arr(i)
            D_XLookupMany = return_elem
            Exit Function
        End If
        
        
    Next
'turn lookup_arr to 2d
'turn return_arr to 1d
'turn value to 1d
'look through value by index
'If the
    
End Function
Function D_XLookupMany2(value_arr, lookup_arr, return_arr)
'value_arr is in order alianed with lookup_arr

End Function
Function D_XlookUp(lookup_value, lookup_array, return_array)

End Function

Function D_Xlookup2D(table, my_row, my_col)
'@@@@@@@@@@@@@@@@Dependency -> Space&Select
'Sp_PickColumn, Sp_PickRow
    col_name = Sp_PickRow(table, 1)
    row_name = Sp_PickColumn(table, 1)
    return_row = VB_Xlookup(my_row, row_name, table)
    res = VB_Xlookup(my_col, col_name, return_row)
    D_Xlookup2D = res
    

End Function
