Attribute VB_Name = "Func_Array2"
Function A_Count(arr As Variant, Optional elements = "") As Long
' Count the number of elements that are actually in the array
'I have to create this because sometimes array index start with 0,1
'It's hard to count the actual elements using only Ubound
'That's why I'm creating this

'This function works by counting elements(or array of elements) in the arrays
'If element is set to Nothing this function will count all the elements in array

    If A_NDim(elements) = 2 Then
        elements = A_Reshape_2dTo1D(elements)
    End If
    
    If Not IsArray(elements) Then
    'Use this If to prevent Error when compare array with ""
        If elements = "" Then
            If IsArray(arr) Then
                A_Count = UBound(arr) - LBound(arr) + 1
            Else
                A_Count = 1
            End If
            Exit Function
        End If
    End If
    
    Dim count As Long
    count = 0
    
    If A_NDim(arr) = 1 Then
        If IsArray(elements) Then
            For i = LBound(arr) To UBound(arr)
                For j = LBound(elements) To UBound(elements)
                    If arr(i) = elements(j) Then
                        count = count + 1
                        Exit For
                    End If
                Next j
            Next i
        Else
            For i = LBound(arr) To UBound(arr)
                If arr(i) = elements Then
                    count = count + 1
                End If
            Next i
        End If

    ElseIf A_NDim(arr) = 2 Then
        If IsArray(elements) Then
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    For k = LBound(elements) To UBound(elements)
                        If arr(i, j) = elements(k) Then
                            count = count + 1
                            Exit For
                        End If
                    Next k
                Next j
            Next i
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    If arr(i, j) = elements Then
                        count = count + 1
                    End If
                Next j
            Next i
        End If
    Else
        MsgBox ("Can't handle array more than 2 dim")
    End If
    A_Count = count


    
End Function

Function A_ShiftIndex(arr, shift)
    Dim outArray() As Variant
    i_min = LBound(arr, 1) + shift
    i_max = UBound(arr, 1) + shift
    j_min = LBound(arr, 2) + shift
    j_max = UBound(arr, 2) + shift
    ReDim outArray(i_min To i_max, j_min To j_max)
'Not done add case that is 1d
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            myVal = arr(i, j)
            outArray(i + shift, j + shift) = myVal
        Next
    Next
    A_ShiftIndex = outArray

'right now assume that arr is 2d
End Function
Function A_Append2(old_arr, new_elem)
'Hard for ChatGPT
'Not Done if done merge with A_Append
'Work not to sure that it's good enough to merge
'Need more testing
    Dim outArr() As Variant
    On Error GoTo ErrorTask
    checker = IsError(UBound(old_arr))
    'Shift Index all to start with 0
    If LBound(old_arr) = 1 Then
        old_arr = A_ShiftIndex(old_arr, -1)
    End If
    If IsArray(new_elem) Then
        If LBound(new_elem) = 1 Then
            old_arr = A_ShiftIndex(new_elem, -1)
        End If
    End If
    
    
    If checker Then
ErrorTask:
        If IsArray(new_elem) Then
            ReDim Preserve outArr(0, UBound(new_elem))
            For j = 0 To UBound(new_elem)
                outArr(0, j) = new_elem(j)
            Next j
        Else
            ReDim outArr(0)
            outArr(0) = new_elem
        End If
        A_Append2 = outArr
        Exit Function
    End If
'Assume that new_elem is 1d array
'Shift index in old_arr all start with 0
    
    
    
    If IsArray(new_elem) Then
    'In case old_arr is empty
        If A_NDim(old_arr) = 2 Then
            ReDim Preserve outArr(UBound(old_arr, 1) + 1, UBound(old_arr, 2))
            For i = LBound(old_arr, 1) To UBound(old_arr, 1)
                For j = LBound(old_arr, 2) To UBound(old_arr, 2)
                    outArr(i, j) = old_arr(i, j)
                Next
            Next
            
            For j = 0 To UBound(old_arr, 2)
                outArr(UBound(old_arr, 1) + 1, j) = new_elem(j)
            Next j
        
        ElseIf A_NDim(old_arr) = 1 Then
        
            n = UBound(old_arr)
            ReDim Preserve old_arr(n + 1)
            old_arr(n + 1) = new_elem
            outArr = old_arr
        Else
            MsgBox ("Not support array > 2 dim")
        End If
        
    Else
        n = UBound(old_arr)
        ReDim Preserve old_arr(n + 1)
        old_arr(n + 1) = new_elem
        outArr = old_arr
    End If
    
    A_Append2 = outArr
End Function


Function A_NDim(v As Variant) As Long
'Returns number of dimensions of an array or 0 for
'an undimensioned array or -1 if no array at all.
'https://answers.microsoft.com/en-us/msoffice/forum/all/how-many-vba-array-dimensions/a4c80919-3cd3-4ed0-a173-b9b8fabd3c83
    Dim i As Long
    A_NDim = 0
    If Not IsArray(v) Then
        Exit Function
    End If
    On Error GoTo Err01
    i = 1
    Do While True
        checker = UBound(v, i)
        If Err.Number <> 0 Then Exit Do
        i = i + 1
    Loop
Err01:
    A_NDim = i - 1
End Function
Function A_Flip(arr As Variant) As Variant
    Dim i As Long
    Dim result() As Variant
    
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        result(i) = arr(UBound(arr) - i)
    Next i
    
    A_Flip = result
End Function

Function A_ShiftRight(arr As Variant, Optional num_move = 1) As Variant
'Chat GPT Struggle!
''Include shiftleft by specify num_move negative

    Dim i As Long
    Dim result() As Variant
    n = UBound(arr) - LBound(arr) + 1
    If num_move >= 0 Then
        ReDim result(LBound(arr) To UBound(arr))
        
        For i = LBound(arr) To n - num_move - 1
            result(i + num_move) = arr(i)
        Next i
        
        For i = LBound(arr) To num_move - 1
            result(i) = arr(n - num_move + i)
        Next i
        
    Else
        ReDim result(LBound(arr) To UBound(arr))
        
        For i = LBound(arr) - num_move To UBound(arr)
            result(i + num_move) = arr(i)
        Next i
        
        For i = LBound(arr) To LBound(arr) - num_move - 1
            result(n + num_move + i) = arr(i)
        Next i
    End If
    A_ShiftRight = result
End Function





       





Function A_GetBack(arr As Variant, num As Long) As Variant
'From ChatGPT
    Dim result() As Variant
    Dim i As Long
    
    ReDim result(0 To num - 1)
    
    ' Copy the last num elements from the array to the result
    For i = 0 To num - 1
        result(i) = arr(UBound(arr) - num + 1 + i)
    Next i
    
    A_GetBack = result
End Function

Function A_GetFront(arr As Variant, num As Long) As Variant
'From ChatGPT
    Dim result() As Variant
    Dim i As Long
    
    ReDim result(0 To num - 1)
    
    ' Copy the first num elements from the array to the result
    For i = 0 To num - 1
        result(i) = arr(i)
    Next i
    
    A_GetFront = result
End Function

Function A_DeleteBack(arr As Variant, num As Long) As Variant
''From ChatGPT
    Dim result() As Variant
    Dim i As Long
    
    ReDim result(0 To UBound(arr) - num)
    
    ' Copy elements from the array to the result
    For i = 0 To UBound(result)
        result(i) = arr(i)
    Next i
    
    A_DeleteBack = result
End Function
Function A_DeleteFront(arr As Variant, num As Long) As Variant
'From ChatGPT
    Dim result() As Variant
    Dim i As Long, j As Long
    
    ReDim result(0 To UBound(arr) - num)
    
    ' Shift elements in the array to the left
    For i = num To UBound(arr)
        result(j) = arr(i)
        j = j + 1
    Next i
    
    A_DeleteFront = result
End Function

Function A_1dTo2d(arr_in)
    Dim out_arr() As Variant
    n = UBound(arr_in)
    ReDim out_arr(n, 0)
    For i = 0 To n
        out_arr(i, 0) = arr_in(i)
    Next
    A_1dTo2d = out_arr
    
End Function
Function A_IsArray(input_in)
    If TypeOf input_in Is range Then
        A_IsArray = False
        Exit Function
    Else
        If IsArray(input_in) Then
            A_IsArray = True
            Exit Function
        End If
    
    End If
End Function

