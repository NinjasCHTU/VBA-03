Attribute VB_Name = "Func_SRand1"
Function R_pick(seed, in_range, Optional n = 1)
    my_arr = A_toArray1d(in_range)
    Randomize seed
    
    a = LBound(my_arr)
    b = UBound(my_arr)
    
    random_inx = Int((b - a + 1) * Rnd + a)
    out_elem = my_arr(random_inx)
    R_pick = out_elem
End Function

Function R_PseudoUnit(x)
    '@@@@@@@@@@@@@@@@Dependency -> no
    a = 75: c = 74: m = 65537
    res = (a * x + c) Mod m
    R_PseudoUnit = res
End Function

Function R_Start_srand(seed, jump)
    '@@@@@@@@@@@@@@@@Dependency -> no
    temp = seed
    If jump = 0 Then
        R_Start_srand = seed
    End If
    
    For i = 0 To jump - 1
        temp = R_PseudoUnit(temp)
    Next i
    R_Start_srand = temp
End Function

Function R_srand_Hrz(seed, Optional n = 1)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim arr_Num() As Variant
    Dim arr_deci() As Variant
    m = 65537
    ReDim arr_Num(n - 1)
    ReDim arr_deci(n - 1)
    
    start_val = R_Start_srand(seed, 2)
    arr_Num(0) = start_val
    For i = 1 To n - 1
        arr_Num(i) = R_PseudoUnit(arr_Num(i - 1))
    Next i
    
    For i = 0 To n - 1
        arr_deci(i) = arr_Num(i) / m
    Next i
    R_srand_Hrz = arr_deci
    
End Function

Function R_srand_int_Hrz(seed, a, b, Optional ByVal n As Variant = 1)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim arr_Num() As Variant
    Dim arr_int() As Variant
    m = 65537
    ReDim arr_Num(n - 1)
    ReDim arr_int(n - 1)
    
    start_val = R_Start_srand(seed, 2)
    arr_Num(0) = start_val
    For i = 1 To n - 1
        arr_Num(i) = R_PseudoUnit(arr_Num(i - 1))
    Next i
    d = b - a + 1
    
    For i = 0 To n - 1
        arr_int(i) = (arr_Num(i) Mod d) + a
    Next i
    R_srand_int_Hrz = arr_int
End Function

Function R_srand_Int(seed, a, b, Optional n = 1, Optional output_type = 0)
    'D_OutputFormat @Dear1
    Dim outArr() As Variant
    
    For i = 1 To n
        random_no = Int((b - a + 1) * Rnd + a)
        outArr = A_Append(outArr, random_no)
    Next
    
    outArr = D_OutputFormat(outArr, output)
    R_srand_Int = outArr
    
End Function

Function R_srand_deci_Hrz(seed, a, b, Optional ByVal n As Variant = 1)
'@@@@@@@@@@@@@@@@Dependency -> no
    temp_arr = R_srand_Hrz(seed, n)
    Dim arr_deci() As Variant
    ReDim arr_deci(n - 1)
    d = b - a
    For i = 0 To UBound(temp_arr)
        arr_deci(i) = d * temp_arr(i) + a
    Next i
    R_srand_deci_Hrz = arr_deci

End Function

Function R_srand(seed, Optional n = 1)
    '@@@@@@@@@@@@@@@@Dependency ->
'Lib_WSFunc(VB_Transpose)
    temp_arr = R_srand_Hrz(seed, n)
    R_srand = VB_Transpose(temp_arr)
End Function




Function R_srand_deci(seed, a, b, Optional ByVal n As Variant = 1)
    '@@@@@@@@@@@@@@@@Dependency ->
'Lib_WSFunc(VB_Transpose)
    temp_arr = R_srand_deci_Hrz(seed, a, b, n)
    R_srand_deci = VB_Transpose(temp_arr)
End Function


'Need Fixed
Function R_srand_rank(seed, n)
    '@@@@@@@@@@@@@@@@Dependency ->
'Lib_WSFunc(VB_Seq&VB_SortBy)
    
    a = VB_Seq(n)
    b = R_srand_Hrz(seed, n)
    res = VB_SortBy(VB_Seq(n), R_srand_Hrz(seed, n))
    R_srand_rank = res
End Function




