Attribute VB_Name = "Work05_Rotation2"
Function Rg_TopORBottom(rng01, rng02)
    If rng01.row < rng02.row Then
        Rg_TopORBottom = "top"
    ElseIf rng01.row > rng02.row Then
        Rg_TopORBottom = "bottom"
    Else
        Rg_TopORBottom = "same row"
    End If
End Function

Function Rg_LeftORRight(rng01, rng02)
'assume rng01, rng02 to be 1 cell
    If rng01.column < rng02.column Then
        Rg_LeftORRight = "left"
    ElseIf rng01.column > rng02.column Then
        Rg_LeftORRight = "right"
    Else
        Rg_LeftORRight = "same column"
    End If
End Function
Function Rg_PickTopOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are above ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refRow = ref_rng_in.row
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refRow = ref_rng.row
    End If
    
    If include = 1 Then
      refRow = ref_rng.row + 1
    ElseIf include = 0 Then
      refRow = ref_rng.row
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
        curr_row = rng.row
        If curr_row < refRow Then
            Set output_range = Rg_Union(output_range, rng)
        End If
    Next
    
    Set Rg_PickTopOf = output_range
End Function

Function Rg_PickBottomOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are below ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refRow = ref_rng_in.row
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refRow = ref_rng.row
    End If
    
    If include = 1 Then
      refRow = ref_rng.row - 1
    ElseIf include = 0 Then
      refRow = ref_rng.row
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
        curr_row = rng.row
        If curr_row > refRow Then
            Set output_range = Rg_Union(output_range, rng)
        End If
    Next
    Set Rg_PickBottomOf = output_range
End Function

Function Rg_PickLeftOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are left ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refCol = ref_rng_in.column
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refCol = ref_rng.column
    End If
    
    If include = 1 Then
      refCol = ref_rng.column + 1
    ElseIf include = 0 Then
      refCol = ref_rng.column
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
    curr_col = rng.column
    If curr_col < refCol Then
        Set output_range = Rg_Union(output_range, rng)
    End If
    Next
    Set Rg_PickLeftOf = output_range
End Function
Function Rg_PickRightOf(rngs_in, ref_rng_in, Optional include = 1)
'include = 1: include the ref_rng row
'include = 0: Not include the ref_rng row
'Work with both range and string
' This function filter out rnsg_in and left with the ranges that are right ref_rng_in
    Dim rng As range
    Dim output_range As range
    
    If TypeName(rngs_in) = "Range" Then
        ws_name = rngs_in.Parent.name
        Set rngs = rngs_in
    Else
        ws_name = ActiveSheet.name
        Set ws = Worksheets(ws_name)
        Set rngs = ws.range(rngs_in)
    End If

    If TypeName(ref_rng_in) = "Range" Then
        refCol = ref_rng_in.column
        Set ref_rng = ref_rng_in
    Else
        Set ref_rng = ws.range(ref_rng_in)
        refCol = ref_rng.column
    End If
    
    If include = 1 Then
      refCol = ref_rng.column - 1
    ElseIf include = 0 Then
      refCol = ref_rng.column
    Else
      Debug.Print ("Enter the valid include value(1 or 0)")
    End If

    For Each rng In rngs
    curr_col = rng.column
    If curr_col > refCol Then
        Set output_range = Rg_Union(output_range, rng)
    End If
    Next
    Set Rg_PickRightOf = output_range
End Function


