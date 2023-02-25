Attribute VB_Name = "Func_String1"
Function St_Repeat(elem_list, n, Optional output_option = 0)
'''' FIX THIS it does't repeat
      '@@@@@@@@@@@@@@@@Array -> LibArray1
       '@@@@@@@@@@@@@@@@Array -> WSFunc
    Dim outArray() As Variant
    Dim outArray2() As Variant
    
    For Each curr_cell In elem_list
        curr_elem = curr_cell.value
        arr_ori = A_Append(outArray, curr_elem)
    Next
    
    
    outArray = arr_ori
    If n > 1 Then
        For i = 2 To n
            outArray = A_Extend(outArray, arr_ori)
        Next
    End If
        
    
    Select Case output_option
    Case 1
        St_Repeat = outArray
    Case 0
        St_Repeat = VB_Transpose(outArray)
    End Select
    
    
End Function
Function St_StrReverse(str_in As String)
    St_StrReverse = StrReverse(str_in)
End Function
Function St_InStrMany(string_in As String, pattern As String, Optional mode_in = vbTextCompare)
  '@@@@@@@@@@@@@@@@Array -> LibArray1
    curr_inx = 1
    Dim out_arr() As Variant
    
    Do
        curr_inx = InStr(curr_inx, string_in, pattern, mode_in)
    
        
        If curr_inx <> 0 Then
            out_arr = A_Append(out_arr, curr_inx)
        End If
        
    Loop While curr_inx > 0
    St_InStrMany = out_arr
    
End Function

'Can expand Function to make it count many characters at once
Function St_count(inString, ch)
  '@@@@@@@@@@@@@@@@Dependency -> No
    n = Len(inString)
    n_ch = Len(ch)
    Count = 0
    For i = 1 To n
        curr_string = Mid(inString.value, i, n_ch)
        If curr_string = ch Then
            Count = Count + 1
        End If
    Next
    St_count = Count
End Function

Function St_CutRightVB(text, n)
  '@@@@@@@@@@@@@@@@Dependency -> No
    St_CutRightVB = Left(text, Len(text) - n)
End Function

Function St_CutLeftVB(text, n)
  '@@@@@@@@@@@@@@@@Dependency -> No
    St_CutLeftVB = Right(text, Len(text) - n)
End Function
Function St_TxtByInxVB(text, start_inx, end_inx)
  '@@@@@@@@@@@@@@@@Dependency -> No
    St_TxtByInxVB = Mid(text, start_inx, end_inx - start_inx + 1)
End Function

Function S_Remove1(inString, ch)

End Function

Function St_RemoveAll(inString, ch)
  '@@@@@@@@@@@@@@@@Dependency -> No
    str02 = Replace(inString, ch, "")
    St_RemoveAll = str02
End Function

'How St_ReplaceBy different from Substitute???
Function St_ReplaceBy(text, new_text As String, old_text_arr As Variant)
  '@@@@@@@@@@@@@@@@Dependency -> No
    'old_text_arr = array that contains old alphabets for the replacement
    n_text = Len(text)
    n_arr = UBound(old_text_arr)

    out_str = ""
    
    For i = 1 To n_text
        curr_ch = Mid(text, i, 1)
        For j = LBound(old_text_arr) To UBound(old_text_arr)
            If (curr_ch = old_text_arr(j)) Then
                curr_ch = new_text
            End If

        Next j
        out_str = out_str & curr_ch
    Next i
    St_ReplaceBy = out_str
End Function

Function St_UnDiaCriticVB(ch)
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim a_varyForm, e_varyForm, i_varyForm, o_varyForm, u_varyForm, y_varyForm, c_varyForm, n_varyForm As Variant
    a_varyForm = Array("ä", "á", "â", "à", "å", "ã")
    e_varyForm = Array("e", "ë", "é", "ê", "è")
    i_varyForm = Array("ï", "í", "î", "ì")
    o_varyForm = Array("è", "õ", "ô", "ò", "ó")
    u_varyForm = Array("ü", "ú", "û", "ù")
    y_varyForm = Array("ÿ")
    c_varyForm = Array("ç")
    n_varyForm = Array("ñ")
    
    un_a_str = St_ReplaceBy(ch, "a", a_varyForm)
    un_ae_str = St_ReplaceBy(un_a_str, "e", e_varyForm)
    un_aei_str = St_ReplaceBy(un_ae_str, "i", i_varyForm)
    un_aeo_str = St_ReplaceBy(un_aei_str, "o", o_varyForm)
    un_aeou_str = St_ReplaceBy(un_aeo_str, "u", u_varyForm)
    un_aeouy_str = St_ReplaceBy(un_aeou_str, "y", y_varyForm)
    un_aeouyc_str = St_ReplaceBy(un_aeouy_str, "c", c_varyForm)
    un_aeouycn_str = St_ReplaceBy(un_aeouyc_str, "n", n_varyForm)
    
    
    final_string = un_aeouycn_str
    St_UnDiaCriticVB = final_string
    
    
End Function
