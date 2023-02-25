Attribute VB_Name = "Func_Dear1"
'All Lib Needed
'
Function D_TextFormula(text_in As String)
    n = Len(text_in)
    i = 1
    out_text = "=" & Chr(34)
    Do While i <= n
        j = 1
        curr_alpha = Mid(text_in, i, j)
        i_old = i
        check02 = "Nope"
        check = "NotEnter"
        If curr_alpha = "V" Then
            temp = "For debug stop point"
        End If
        Do While IsNumeric(curr_alpha) And InStr(1, curr_alpha, ",") = 0
            curr_alpha = Mid(text_in, i_old, j)
            j = j + 1
            i = i + 1
            check = "Enter"
            If i > n Then
                check02 = "EndWithNum"
                Exit Do
            End If
        Loop
        If check <> "Enter" Then
            i = i + 1
            out_text = out_text & curr_alpha
        Else
            If check02 = "EndWithNum" Then
                number_str = curr_alpha
                out_text = out_text & Chr(34) & "   &" & number_str
            Else
                number_str = St_CutRightVB(curr_alpha, 1)
                ch01 = Right(curr_alpha, 1)
                out_text = out_text & Chr(34) & "   &" & number_str & "&   " & Chr(34) & ch01
            End If
            
        End If
    Loop
    If check02 <> "EndWithNum" Then
        out_text = out_text & Chr(34)
    End If
    D_TextFormula = out_text
End Function

Function D_OutputFormat(arr_in, output_option)
    Select Case output_option
    Case 0
        D_OutputFormat = VB_Transpose(arr_in)
    Case 1
        D_OutputFormat = arr_in
    Case Else
        MsgBox ("Enter the type only:   0:Vertical(Normal Use), 1:Horizontal(Used in VBA) ")
        D_OutputFormat = "Invalid Type"
    End Select
End Function


Function D_isItInWB(word_in, Optional output_option = 0, Optional transpose_option = 0)
'output_option = 0  Both sheet_Name&Address
'output_option = 1  only sheetname

'transpose_option = 0 Vertical
'transpose_option = 1 Horizontal (as Array)

'@@@@@@@@@@@@@@@@Dependency -> Array1
'WS_Func => Transpose
'Not Done Next: Try to remove the referred text addresss
'1) Add Option to show only SheetName/ address Only
'2) Add Option for tranpose
    
'To keep address & sheet_name seperately
    Dim address_arr() As Variant
    Dim sheet_arr() As Variant
    Dim address_and_sheet_arr() As Variant

    mySheetNameList = O_GetSheetName(1)
    word_address = word_in.Address
    
    For i = LBound(mySheetNameList) To UBound(mySheetNameList)
        curr_SheetName = mySheetNameList(i)
        curr_addr_sheet = D_isItInWS(word_in, curr_SheetName, 0)
'Reshape it to 1D
        
        If TypeName(curr_addr_sheet) = "String" Then
            
        Else
'2D Case
            On Error GoTo Err01_2DCase
Back01:
            sheet_arr = A_Append(sheet_arr, curr_SheetName)
            address_and_sheet_arr = A_Union(address_and_sheet_arr, curr_addr_sheet)
        End If

    Next i
    Select Case output_option
    'output_option = 0  Both sheet_Name&Address
        Case 0
            D_isItInWB = D_OutputFormat(address_and_sheet_arr, transpose_option)
    'output_option = 1  only sheetname
        Case 1
            D_isItInWB = D_OutputFormat(sheet_arr, transpose_option)
        
    End Select
    Exit Function
Err01_2DCase:
    curr_addr_sheet = A_Reshape_2dTo1D(curr_addr_sheet)
    GoTo Back01
    
    


End Function

Function D_isItInWS(word_in, mySheetName, Optional output_option = 2, Optional transpose_option = 0)
'

'output_option = 0: Both SheetName&Address
'output_option = 1: Only SheetName
'output_option = 2: Only Address

'transpose_option = 0 Vertical
'transpose_option = 1 Horizontal (as Array)

'@@@@@@@@@@@@@@@@Dependency -> Array1
'WS_Func => Transpose
'Not Done Next: Try to remove the referred text addresss
'1)!!!!If input is string and not Range
'2)
    Dim myWS As Worksheet
    Dim resFind0 As range
    Dim searchRange As range
    Dim curr_found As range
    Dim out_address() As Variant
    Dim out_arr() As Variant
'To keep address & sheet_name seperately
    Dim address_arr() As Variant
    Dim sheet_name_arr() As Variant

    mySheetName = mySheetName
    word_address = word_in.Address
    
    
    Set searchRange = Sheets(mySheetName).UsedRange
    Set myWS = Sheets(mySheetName)
    Set resFind0 = searchRange.Find(what:=word_in, LookIn:=xlValues, lookat:=xlWhole)
    Set curr_found = resFind0
    If resFind0 Is Nothing Then
        D_isItInWS = "Not Found"
        Exit Function
    End If
    first_address = resFind0.Address
    
    Do While True
        Set curr_found = searchRange.Find(what:=word_in, After:=curr_found, LookIn:=xlValues, lookat:=xlWhole)
        found_sheet_name = curr_found.Parent.name
        found_address = curr_found.Address
        saved_address = found_sheet_name & "   " & found_address


        If found_address <> word_address Then
            out_arr = A_Append(out_arr, saved_address)
            address_arr = A_Append(address_arr, found_address)
            sheet_name_arr = A_Append(sheet_name_arr, found_sheet_name)
        End If
        If curr_found.Address = first_address Then
            Exit Do
        End If
       
    Loop
    
    'addressFound = resFind.Address
    'D_isItInWS = addressFound
    If A_isEmpty(out_arr) Then
        D_isItInWS = "Not Found"
    Else
        If transpose_option = 1 Then
            D_isItInWS = out_arr
        Else
            Select Case output_option
            Case 0
                D_isItInWS = VB_Transpose(out_arr)
            Case 1
                D_isItInWS = VB_Transpose(sheet_name_arr)
            Case 2
                D_isItInWS = VB_Transpose(address_arr)
            End Select
            
        End If
    End If
End Function

Function D_isItInThisWS(word_in2, Optional output_option = 0)
    Dim this_sheet_name As String
    this_sheet_name = Application.Caller.Parent.name
    
    res = D_isItInWS(word_in2, this_sheet_name)
    D_isItInThisWS = res

End Function
''''!!!!!! THE Problem happens when I18 is searched it refers to itself
'''' !!!!! So oneway to fix this is know that way to remove particular cell/or (more general range)
'''' !!!! From the range that we have
Function D_Rank(elem, look_col, value_col, Optional option_1 = 0)
Attribute D_Rank.VB_Description = "Find the ranking of the element from the value list"
Attribute D_Rank.VB_ProcData.VB_Invoke_Func = " \n14"
    myValue = Application.XLookup(elem, look_col, value_col)
    res = WorksheetFunction.Rank_Avg(myValue, value_col, option_1)
    D_Rank = res
    
    
    
End Function



Function D_isItInRange(rangeIn, elem)
'@@@@@@@@@@@@@@@@Dependency -> no
    res = False
    For Each curr_cell In rangeIn
        curr_adress = curr_cell.Address
        If curr_cell.value = elem Then
            res = curr_cell.Address
        End If
        
        
    Next curr_cell
    D_isItInRange = res
    

End Function
Function D_AlphaSmallV(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(97 + i)
    Next i
    D_AlphaSmallV = WorksheetFunction.Transpose(arr)
End Function

Function D_AlphaSmallH(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(97 + i)
    Next i
    D_AlphaSmallH = arr

End Function

Function D_AlphaBigV(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(65 + i)
    Next i
    D_AlphaBigV = WorksheetFunction.Transpose(arr)

End Function
Function D_AlphaBigH(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(65 + i)
    Next i
    D_AlphaBigH = (arr)

End Function
Function D_Combi2(arr1, arr2)
'the output is array(prefereed)
    arr2 = A_toArray1d(arr2)
    Dim out_arr() As Variant
    
    If TypeOf arr1 Is range Then
        arr1 = A_toArray1d(arr1)
    'If arr1 is 1 dimestion then
    'Just combine them with arr2(normal&easy)
        n1 = UBound(arr1)
        n2 = UBound(arr2)
        n_row = (n1 + 1) * (n2 + 1)
        ReDim out_arr(n_row - 1, 1)
        k = 0
        For i = 0 To n1
            curr_alpha1 = arr1(i)
            For j = 0 To n2
                
                If k >= n_row Then
                
                    Exit For
                End If
                out_arr(k, 0) = arr1(i)
                out_arr(k, 1) = arr2(j)
                k = k + 1
                
            Next j

        Next i
        D_Combi2 = out_arr
        Exit Function
        
    Else
        n1 = UBound(arr1, 1)
        n2 = UBound(arr2)
        left_part = A_VStackRep(arr1, n2 + 1)
        right_part = A_VRepArr(arr2, n1 + 1)
'Still got the error because right_part is 1d but A_HStack works only for 2d
        right_part = A_1dTo2d(right_part)
        out_arr = A_HStack(left_part, right_part)
        D_Combi2 = out_arr
        Exit Function
    
    End If
    

    
    'arr1 can has 1 dimension or 2 dimenstion

    
    'If arr1 has 2 dimesions then
    'Create another column
    'Loop through arr2 fill in element of arr2 til end
    
    
End Function
Function D_Combination(ParamArray range_tab())
    Dim out_arr() As Variant
    n_row = 1
    n_col = UBound(range_tab)
    temp_arr = D_Combi2(range_tab(0), range_tab(1))
    For i = 2 To UBound(range_tab)
        temp_arr = D_Combi2(temp_arr, range_tab(i))
    Next
    D_Combination = temp_arr
    'Use D_Combi2 over and over again
    
    
    
End Function
' Function D_Combination(ParamArray range_tab())
' '''''''''''''''''''''''''''''''''' old version '''''''''''''''''''''''''''''''''''
' ' FIXME Need fixing the index are not complete at the end
'     '@@@@@@@@@@@@@@@@Dependency ->
' ' Lib_WSFunc(VB_Transpose)
'     temp_tab = D_Combi2(range_tab(0), range_tab(1), 3)
    
'     For i = 2 To UBound(range_tab)
'         temp_tab = D_Combi2(temp_tab, range_tab(i), 3)
'     Next i
    
'     D_Combination = VB_Transpose(temp_tab)
' End Function

' Function D_Combi2(arr1, arr2, Optional choice = 2)
' '''''''''''''''''''''''''''''''''' old version '''''''''''''''''''''''''''''''''''
' '@@@@@@@@@@@@@@@@Dependency ->
' ' Lib_WSFunc(VB_Transpose)
'     If TypeOf arr1 Is Range Then
'         row_n = arr1.Count
'     Else
'         row_n = UBound(arr1)
'     End If
    
'     If TypeOf arr2 Is Range Then
'         col_n = arr2.Count
'     Else
'         col_n = UBound(arr2)
'     End If
    
    
'     Dim arr2d(3, 2) As Variant
    
'     arr2d_new = D_2dArray(arr2d, row_n, col_n)
'     Dim arr1d() As Variant
'     ReDim Preserve arr1d(row_n * col_n - 1)
    
'     'k is the index for arr1d
'     k = 0
'     '1&2 is the dimention
'     For i = 0 To UBound(arr2d_new, 1)
'         For j = 0 To UBound(arr2d_new, 2)
'             If TypeOf arr1 Is Range Then
'                 new_elem = arr1(i + 1) & " " & arr2(j + 1)
'             Else
'                 new_elem = arr1(i) & " " & arr2(j + 1)
'             End If
            
'             arr2d_new(i, j) = new_elem
'             arr1d(k) = new_elem
'             k = k + 1
'         Next j
'     Next i
    
'     If choice = 1 Then
'         D_Combi2 = VB_Transpose(arr1d)
'     ElseIf choice = 2 Then
'         D_Combi2 = arr2d_new
'     Else
'         D_Combi2 = arr1d
'     End If
    
    
' End Function

Function D_2dArray(name, row, col)
  '@@@@@@@@@@@@@@@@Dependency -> No
    D_2dArray = ReDimPreserve(name, row - 1, col - 1)
End Function


Function D_sheetTOArr(table)
'Not Done
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr(11) As Variant
    row = table.Rows.Count
    col = table.Columns.Count
    arr02 = D_2dArray(arr, row, col)
    i = 0
    For Each c In table
        arr(i) = c
        i = i + 1
    Next c
    D_sheetTOArr = arr
End Function
Function make_key_num()
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim key_num As Object
    Set key_num = CreateObject("Scripting.Dictionary")
    key_num("C") = 1
    key_num("C#") = 2
    key_num("D") = 3
    key_num("D#") = 4
    key_num("E") = 5
    key_num("F") = 6
    key_num("F#") = 7
    key_num("G") = 8
    key_num("Ab") = 9
    key_num("A") = 10
    key_num("Bb") = 11
    key_num("B") = 12
    
    key_num("Db") = 2
    key_num("Eb") = 4
    key_num("Gb") = 7
    key_num("G#") = 9
    key_num("A#") = 11
    Set make_key_num = key_num

End Function
Function make_num_2_key()
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim num_2_key() As Variant
    num_2_key = Array("C", "C#", "D", "D#", "E", "F", "F#", "G", "Ab", "A", "Bb", "B")
    make_num_2_key = num_2_key

End Function
Function M_TransposeVB(chord, ori_key As String, target_key As String)
'Change key by telling the original key and target key
'Instead of I mentally chosing the number
    Dim key_num As Object
    Set key_num = CreateObject("Scripting.Dictionary")
    num_2_key = make_num_2_key
    Set key_num = make_key_num
    
    ori_num = key_num(ori_key)
    target_num = key_num(target_key)
    shift = target_num - ori_num
    new_chord = M_ChangeKeyVB(chord, shift)
    M_TransposeVB = new_chord
    'M_ChangeKey2 = new_val
   

End Function
Function M_ChangeKey2(chord, shift)
      '@@@@@@@@@@@@@@@@Dependency -> No
    Dim key_num As Object
    Set key_num = CreateObject("Scripting.Dictionary")
    
    num_2_key = make_num_2_key
    Set key_num = make_key_num
    
    If TypeOf chord Is range Then
        old_val = key_num(chord.value)
    Else
        old_val = key_num(chord)
    
    End If
    
    new_val = (old_val + shift) Mod 12
    
    If new_val <= 0 Then
        new_val = 12 + new_val
    End If
    
    'M_ChangeKey2 = new_val
    M_ChangeKey2 = num_2_key(new_val - 1)
    
    'MsgBox (key_num("b"))
    
End Function
Function M_ChangeKeyVB(chord_txt, shift)
'$$$$$$$$$$$$$$$$$$$$$$ I can at the output format
'@@@@@@@@@@@@@@@@Dependency ->
    'Lib_String1
    each_chord = Split(chord_txt, " ")
   
    Dim newChordArr() As Variant
    ReDim newChordArr(UBound(each_chord))

    For i = LBound(each_chord) To UBound(each_chord)
        curr_chord = each_chord(i)
        
        con1 = (InStr(1, curr_chord, "b") > 0)
        con2 = (InStr(1, curr_chord, "#") > 0)
        
        If (InStr(1, curr_chord, "b") > 0) Or (InStr(1, curr_chord, "#") > 0) Then
            curr_key = St_TxtByInxVB(curr_chord, 1, 2)
            curr_add = St_CutLeftVB(curr_chord, 2)
            
        Else
            curr_key = St_TxtByInxVB(curr_chord, 1, 1)
            curr_add = St_CutLeftVB(curr_chord, 1)
        
        End If
        new_key = M_ChangeKey2(curr_key, shift)
        new_chord = new_key & curr_add
        newChordArr(i) = new_chord

    Next i

    
    new_chord_str = Join(newChordArr, " ")
    M_ChangeKeyVB = new_chord_str
    
    
    

End Function

'Add feature to also include number from the worksheet directly
'make it work with Array and Range
Function D_toDict(arr01, arr02)
    Dim return_dict As New Scripting.Dictionary
    'If Typeof(arr01) is Array then
    
    
End Function




Sub Test03()
    str01 = "[1,2,3,4]"
    str02 = St_RemoveAll(str01, ",")
    MsgBox (str02)
End Sub


