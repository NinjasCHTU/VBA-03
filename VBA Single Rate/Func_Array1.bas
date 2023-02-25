Attribute VB_Name = "Func_Array1"


Function A_VRepArr(arr_in, n)
'Expect arr_in to be 1d array
    n_arr = UBound(arr_in)
    Dim out_arr() As Variant
    ReDim out_arr((n_arr + 1) * n - 1)
    
    k = 0
    For i = 0 To n_arr
        For j = 0 To n - 1
            out_arr(k) = arr_in(i)
            k = k + 1
        Next
    Next
    A_VRepArr = out_arr
End Function
Function A_VStack(arr1, arr2)
    ' arr1 and arr2 is 2d Array
    n_row1 = UBound(arr1, 1)
    n_col1 = UBound(arr1, 2)
    n_row2 = UBound(arr2, 1)
    n_col2 = UBound(arr2, 2)
    If n_col1 <> n_col2 Then
        A_VStack = "The number of column is not equal"
        Exit Function
    End If
    
    Dim out_arr() As Variant
    ReDim out_arr(n_row1 + n_row2 + 1, n_col1)
    For i = 0 To n_row1
        For j = 0 To n_col1
            out_arr(i, j) = arr1(i, j)
        Next
    Next
    For i = 0 To n_row2
        For j = 0 To n_col2
            out_arr(i + n_row1 + 1, j) = arr2(i, j)
        Next
    Next
    A_VStack = out_arr
    

End Function


Function A_VStackRep(arr_in, n)
    temp_arr = arr_in
    For i = 1 To n - 1
        temp_arr = A_VStack(temp_arr, arr_in)
    Next i
    A_VStackRep = temp_arr

End Function
Function A_NumMatrix(n_row, n_col)
    Dim outArr() As Variant
    ReDim outArr(n_row - 1, n_col - 1)
    For i = 1 To n_row
        For j = 1 To n_col
            curr_val = 10 * i + j
            outArr(i - 1, j - 1) = curr_val
        Next
    Next
    A_NumMatrix = outArr
End Function
Function A_Reshape_1dTo2D(arr_in As Variant, n_row As Long, n_col As Long) As Variant()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim arr_out() As Variant
    k = 0
    ReDim arr_out(0 To n_row - 1, 0 To n_col - 1)
    For i = 0 To n_row - 1
        For j = 0 To n_col - 1
            arr_out(i, j) = arr_in(k)
            k = k + 1
        Next j
    Next i
    A_Reshape_1dTo2D = arr_out
End Function
Function A_Reshape_2dTo1D(arr_in)
'Assume that it's a retangle 2d array
    Dim out_arr() As Variant
    n_row = UBound(arr_in)
    n_col = UBound(arr_in, 2)
    For j = LBound(arr_in, 2) To n_col
        For i = LBound(arr_in, 1) To n_row
            curr_elem = arr_in(i, j)
            out_arr = A_Append(out_arr, curr_elem)
        Next
    Next
    A_Reshape_2dTo1D = out_arr
'I should create a general reshape Function
'This is only convert 2d to 1d

End Function
Function A_isSame(arr1, arr2)
'Work with Empty Array
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            A_isSame = False
            Exit Function
        End If
    Next
    A_isSame = True

End Function
Function A_isElemSame(arr1, arr2)

End Function
Function A_GetUnique(arr1)

End Function
Function A_RemoveN(myArr, elem_to_remove, n)
    Dim outArr() As Variant
    check = 0
    For i = LBound(myArr) To UBound(myArr)
        If myArr(i) = elem_to_remove Then
            If check >= n Then
                outArr = A_Append(outArr, myArr(i))
            End If
            check = check + 1
        Else
            outArr = A_Append(outArr, myArr(i))
        End If
        
    Next
    A_RemoveN = outArr
End Function
Function A_Remove1(myArr, elem_to_remove)
    A_Remove1 = A_RemoveN(myArr, elem_to_remove, 1)
End Function
Function A_RemoveAll(myArr, elem_to_remove)
    Dim outArr() As Variant
    
    For i = LBound(myArr) To UBound(myArr)
        If myArr(i) <> elem_to_remove Then
            outArr = A_Append(outArr, myArr(i))
        End If
    Next
    A_RemoveAll = outArr
End Function
'Incase arr is empty Array
Function A_Union(arr1, arr2)
'Work with Empty Array
    If A_IsEmpty(arr1) Then
        A_Union = arr2
        Exit Function
    ElseIf A_IsEmpty(arr2) Then
        A_Union = arr1
        Exit Function
    End If
    For i = LBound(arr2) To UBound(arr2)
        curr_elem = arr2(i)
        If Not A_isInArr(arr1, curr_elem) Then
            arr1 = A_Append(arr1, curr_elem)
        End If
    Next
    A_Union = arr1

End Function

Function A_Intersect(arr1, arr2)
'Work with Empty Array
    Dim outArr() As Variant
    
    If A_IsEmpty(arr1) Then
        A_Intersect = outArr
        Exit Function
    ElseIf A_IsEmpty(arr2) Then
        A_Intersect = outArr
        Exit Function
    End If
    
    
    For i = LBound(arr2) To UBound(arr2)
        curr_elem = arr2(i)
        If A_isInArr(arr1, curr_elem) Then
            outArr = A_Append(outArr, curr_elem)
        End If
    Next
    A_Intersect = outArr
    

End Function

Function A_SetSubtract(arr1, arr2)
'Work with Empty Array
    Dim out_arr() As Variant
    
    If A_IsEmpty(arr1) Then
        A_SetSubtract = outArr
        Exit Function
    ElseIf A_IsEmpty(arr2) Then
        A_SetSubtract = arr1
        Exit Function
    End If

    
    out_arr = arr1
    For i = LBound(arr2) To UBound(arr2)
        out_arr = A_RemoveAll(out_arr, arr2(i))
    Next
    A_SetSubtract = out_arr

End Function
Function A_Duplicate(myArr, n)
    Dim outArr() As Variant
    For i = 1 To n - 1
        outArr = A_CombineArray(outArr, myArr)
    Next
    A_Duplicate = outArr
End Function
Function A_Sort(myArr)
'Not Done
    'res = WorsheetFunction.Sort(myArr)
    'A_Sort = res
End Function

Function A_CombineArray(ParamArray myArrList())
'Work with empty array as well
    Dim out_arr() As Variant
    
    On Error Resume Next
    For Each curr_arr In myArrList
        For Each curr_elem In curr_arr
            out_arr = A_Append(out_arr, curr_elem)
        Next
    Next
    A_CombineArray = out_arr
End Function

Public Sub A_printArr(inArray)
'@@@@@@@@@@@@@@@@Dependency -> No
' print array do not support 2d array FIXME(improve)
    Dim printCell, printCell_Temp As range
    Dim arr01() As Variant
    'arr01 = Array(1, 2, 3, 4, 5, 6, 7)
    'n01 = 7
    
    
    
    'Set printCell_Temp = Range("K15")
    Set printCell = Application.InputBox(Title:="Print Array", Prompt:="Select Range to print out", Type:=8)
    If A_IsEmpty(inArray) Then
        GoTo Err01_emptyArray
    End If
    
    On Error GoTo Err01_emptyArray
    n = UBound(inArray)
    For i = 0 To n
        Cells(printCell.row + i, printCell.column) = inArray(i)
    Next i
    Exit Sub
Err01_emptyArray:
    Cells(printCell.row, printCell.column) = "Empty Array"
    
    
    
End Sub
Function A_toArray1dVB(selectRange)
    A_toArray1dVB = VB_Transpose(A_toArray1d(selectRange))
    
End Function



Function A_toArray1d(selectRange, Optional direction = "down")
' Direction = down:  Go through down the column first
' Direction = right: Go through the 1row

'@@@@@@@@@@@@@@@@Dependency -> No
    Dim inRange As range
    'Set selectRange = Application.InputBox(Title:="Import value to Arrays", Prompt:="Select Range to import ARRAYS value", Type:=8)
'turn 2d array to 1d array
    If A_IsArray(selectRange) Then
        Exit Function
    End If
    n = selectRange.Cells.count
    Dim outArray() As Variant
    For Each Area In selectRange.Areas
        If direction = "down" Then
            n_row = Area.Rows.count
            n_col = Area.Columns.count
            For j = 1 To n_col
                For i = 1 To n_row
                    curr_val = Area.Cells(i, j).value
                    outArray = A_Append(outArray, curr_val)
                Next
            Next
        ElseIf direction = "right" Then
            n_row = Area.Rows.count
            n_col = Area.Columns.count
            For i = 1 To n_row
                For j = 1 To n_col
                    curr_val = Area.Cells(i, j).value
                    outArray = A_Append(outArray, curr_val)
                Next
            Next
        Else
            A_toArray1d = "Please Enter the correct direction: right or down "
            Exit Function
        End If
    Next
   A_toArray1d = outArray
    
    
End Function

Function A_FindIndex(arr_in, elem)
    Dim out_arr() As Variant
    
    
    For i = LBound(arr_in) To UBound(arr_in)
        If arr_in(i) = elem Then
            curr_inx = i
            out_arr = A_Append(out_arr, curr_inx)
        End If
    Next
    
    A_FindIndex = out_arr
    
End Function


Function A_isInArr(arr_in, checker) As Boolean
'@@@@@@@@@@@@@@@@Dependency -> No
    On Error GoTo Err01
    
    For i = LBound(arr_in) To UBound(arr_in)
        If arr_in(i) = checker Then
            A_isInArr = True
            Exit Function
        End If

    Next i
    
    On Error GoTo 0
    
    A_isInArr = False
    Exit Function
    
Err01:
    A_isInArr = False
    Exit Function
    
    


End Function
'A_toArray2d = 2d Version but not combined with original Function
'must continue
Function A_toArray2d(selectRange)
'@@@@@@@@@@@@@@@@Dependency -> No
'But there is GONNA Be a Problem when it's 1 dimesion when I want to use to other part in VBA
    Dim inRange As range
    'Set selectRange = Application.InputBox(Title:="Import value to Arrays", Prompt:="Select Range to import ARRAYS value", Type:=8)
    n_row = selectRange.Rows.count
    n_col = selectRange.Columns.count
    
    
    'A_toArray2d = n_row & "and" & n_col
    
    Dim outArray() As Variant
    ReDim outArray(n_row - 1, n_col - 1)
    
    'A_toArray2d = selectRange.Cells(2, 4).Value
    'A_toArray2d = outArray
    For i = 1 To n_row
        For j = 1 To n_col
        curr_val = selectRange.Cells(i, j).value
        outArray(i - 1, j - 1) = curr_val
        Next j
    Next i
    A_toArray2d = outArray
    
    
End Function




Function A_Append(old_arr, new_elem)
    On Error GoTo ErrorTask
    checker = IsError(UBound(old_arr))
    If checker Then
ErrorTask:
        ReDim old_arr(0)
        old_arr(0) = new_elem
        A_Append = old_arr
        Exit Function
    End If
    
    n = UBound(old_arr)
    ReDim Preserve old_arr(n + 1)
    old_arr(n + 1) = new_elem
    A_Append = old_arr
End Function

Function A_TxtTO1dArr(inString)
'@@@@@@@@@@@@@@@@Dependency ->
'Lib_Dear1 (St_RemoveAll)
    
    Dim return_arr() As Variant
'declear 1d Array using PYTHON syntax
' Still have problem if I want the input to be INT (right now they are String)!!!!
    str02 = Mid(inString, 2, Len(inString) - 2)
    str03 = Split(str02, ",")
    'This loop is for remove space
    For i = LBound(str03) To UBound(str03)
        str03(i) = Trim(str03(i))
        str03(i) = St_RemoveAll(str03(i), """")
        dot_inx = InStr(1, str03(i), ".")
        If dot_inx = 0 Then
        'There are no .
            curr_elem = CInt(str03(i))
        Else
        'There is a dot
            curr_elem = CDbl(str03(i))
        End If
        return_arr = A_Append(return_arr, curr_elem)
    Next
    
    A_TxtTO1dArr = return_arr
    
End Function

'to be Continue
'Still not dealing with removing " " eg "a"
Function A_TxtTO2dArr(inString)
'@@@@@@@@@@@@@@@@Dependency -> No
' Still have problem if I want the input to be INT (right now they are String)!!!!
    'Assuming retangular array
    'declear 2d Array using PYTHON syntax
    str_noBracket = Mid(inString, 2, Len(inString) - 2)
    str_noSpace = Replace(str_noBracket, " ", "")
    
    text_tab = Split(str_noSpace, "],")
    'to get the number of columns first
    str_removeL = Replace(text_tab(0), "[", "")
    str_removeR = Replace(str_removeL, "]", "")
    each_row_tab = Split(str_removeR, ",")

    col = UBound(each_row_tab) + 1
    row = UBound(text_tab) + 1
    'Declear the 2d out_arr
    Dim out_arr As Variant
    ReDim out_arr(row - 1, col - 1)



    For i = LBound(text_tab) To UBound(text_tab)
        str_removeL = Replace(text_tab(i), "[", "")
        str_removeR = Replace(str_removeL, "]", "")
        each_row_tab = Split(str_removeR, ",")
        

        For j = 0 To col - 1
            out_arr(i, j) = each_row_tab(j)
        Next j

        
    Next i
    
    A_TxtTO2dArr = out_arr
        
        
        
End Function

