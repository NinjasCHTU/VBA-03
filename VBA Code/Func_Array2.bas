Attribute VB_Name = "Func_Array2"
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

