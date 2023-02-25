Attribute VB_Name = "Work05_02_Standardized"
Function D_UniqueRow(ParamArray Ranges() As Variant) As Variant
    Dim rowDict As Object
    Set rowDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(Ranges(0), 1) To UBound(Ranges(0), 1)
        Dim key As String
        key = Join(Application.Index(Ranges, i, 0), ChrW(-1))
        If Not rowDict.Exists(key) Then
            rowDict.Add key, Application.Index(Ranges, i, 0)
        End If
    Next i
    
    Dim result As Variant
    ReDim result(1 To rowDict.count, 1 To UBound(Ranges) + 1)
    i = 1
    For Each key In rowDict.Keys
        result(i, 1) = rowDict(key)(1, 1)
        For j = 2 To UBound(Ranges) + 1
            result(i, j) = rowDict(key)(1, j - 1)
        Next j
        i = i + 1
    Next key
    
    D_UniqueRow = result
End Function


Sub W_BorderSubBulk(start_cell, num_arr)
'Hard for ChatGPT
    Set current_cell = start_cell
    For i = LBound(num_arr) To UBound(num_arr)
        
        If num_arr(i) > 1 Then
            Set temp = current_cell.Resize(num_arr(i), 1)
            Call W_BorderSubElement(temp)
        Else
            Call W_BorderSubElement(current_cell)
        End If
        
        cumm_start = cumm_start + num_arr(i)
        
        Set current_cell = start_cell.Offset(cumm_start, 0)
    Next i
End Sub
Sub W_BorderSubElement(rng)
    Call W_BorderOutside(rng)
    If rng.Rows.count > 1 Then
        Set rng02 = rng.Resize(rng.Rows.count - 1)
        For Each cell In rng02
        ' Do something with the cell
            With cell.Borders(xlEdgeBottom)
                .LineStyle = xlDash
            End With
        Next cell
    End If
End Sub
Sub W_BorderOutside(rng, Optional myLineStyle As XlLineStyle = xlContinuous, Optional myWeight As XlBorderWeight = xlThin)
'Change colorIndex to input hex code as well
    With rng.Borders(xlEdgeLeft)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .colorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .colorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .colorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .colorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = myLineStyle
        .Weight = myWeight
        .colorIndex = xlAutomatic
    End With

End Sub
Function A_AddSpace(arr, num_arr, Optional spaceOption = 0)
' spaceOption = 0 means the total of space
'spaceOption = 1 means # of spaces between 2 cells
'See test_A_AddSpace for example
    Dim outArr() As Variant
    
    For i = LBound(arr) To UBound(arr)
        outArr = A_Append(outArr, arr(i))
        If spaceOption = 0 Then
            If num_arr(i) > 1 Then
                For j = 1 To num_arr(i) - 1
                    outArr = A_Append(outArr, "")
                Next
            End If
        ElseIf spaceOption = 1 Then
            For j = 1 To num_arr(i)
                outArr = A_Append(outArr, "")
            Next
        Else
            A_AddSpace = "Invalid SpaceOption(0 or 1)"
            Exit Function
        End If
        
    Next
    A_AddSpace = outArr

End Function
Sub Rg_Merged(start_cell, num_arr)
'Hard for chatGPT
    Dim i As Long
    Dim current_cell As range
    Dim merged_range As range
    ws_name = start_cell.Parent.name
    Set ws = Sheets(ws_name)
    Set current_cell = ws.range(start_cell.Address)
    cumm_start = 0
    
    For i = LBound(num_arr) To UBound(num_arr)
        
        If num_arr(i) > 1 Then
            Set temp = current_cell.Resize(num_arr(i), 1)
            temp.Merge
        End If
        
        cumm_start = cumm_start + num_arr(i)
        
        Set current_cell = current_cell.Offset(1, 0)
    Next i

    'merged_range.Merge
End Sub

Function Rg_ColFromText(xlWorksheet, text, Optional row_num = 1, Optional output = 0)
'row_num start at 1
'output = 0 => return column as alphabet
'output = 1 => return column as numbers(start with 1)
'Upgrade 01: change ws to accept ws_name
'Upgrade 02: text => change so that it will work with finding substring instead of absolute
    For i = 1 To xlWorksheet.Columns.count
        If xlWorksheet.Cells(row_num, i).value = text Then
            col_inx = i
            colAlphabet = D_AlphaBigHelp01(col_inx)
            Exit For
        End If
    Next i
    If output = 0 Then
        Rg_ColFromText = colAlphabet
    ElseIf output = 1 Then
        Rg_ColFromText = col_inx
    Else
        Rg_ColFromText = "Enter the correct output option(0 or 1)"
    End If

End Function
Function W_GetRedBookData(ws)
    col_make_name = Rg_ColFromText(ws, "Make")
    col_model_name = Rg_ColFromText(ws, "Model")
    col_group_name = Rg_ColFromText(ws, "Group")
    
    add_make = col_make_name & "2"
    add_model = col_model_name & "2"
    add_group = col_group_name & "2"
    
    Set start_make = ws.range(add_make)
    Set start_model = ws.range(add_model)
    Set start_group = ws.range(add_group)
    Dim col_make, col_model, col_group As range
    Set col_make = Rg_PickTilEnd(start_make, xlDown)
    Set col_model = Rg_PickTilEnd(start_model, xlDown)
    Set col_group = Rg_PickTilEnd(start_group, xlDown)
    
    'Set col_make = ws.Columns(col_make_name)
    'Set col_model = ws.Columns(col_model_name)
    'Set col_group = ws.Columns(col_group_name)
    Set combinedCol = Rg_Union(col_make, col_model)
    'Set combinedCol = Union(col_make, col_model, col_group)
    outArr = D_UniqueRow(col_make, col_model, col_group)
    arr_make = col_make.value
    arr_model = col_model.value
    arr_group = col_group.value
    
    'arr_make = A_toArray1d(col_make)
    'arr_model = A_toArray1d(col_model)
    'arr_group = A_toArray1d(col_group)
    
    combined1 = A_HStack(arr_make, arr_model)
    combined2 = A_HStack(combined1, arr_group)
    uniqueCombination = VB_Unique(combinedCol)
    
End Function
Function Rg_GetCellsBetween(str01, str02, ws_name, Optional include = False)
'Assume that only 1 cell exists for each str01,str02
'include = False doesn't include these 2 cells
'include = True include these 2 cells
'Assuming we get the cell from up to down(vertical)
    Set rng01 = Rg_FindAllRange(str01, ws_name)
    Set rng01_Below = rng01.Offset(1, 0)
    Set rng02 = Rg_FindAllRange(str02, ws_name)
    Set rng02_Above = rng02.Offset(-1, 0)
    Set ws = Worksheets(ws_name)
    If include Then
        Set rng = ws.range(rng01, rng02)
    Else
        Set rng = ws.range(rng01_Below, rng02_Above)
    End If

    Set Rg_GetCellsBetween = rng
End Function
Sub S_UnmergedBetween2Cells(str01, str02, ws_name)
'Assume that only 1 cell exists for each str01,str02
    Set rng01 = Rg_FindAllRange(str01, ws_name)
    Set rng02 = Rg_FindAllRange(str02, ws_name)
    Set ws = Worksheets(ws_name)
    ws.range(rng01, rng02).UnMerge
    
End Sub
Function Rg_PickTilEnd(rng, direction As XlDirection) As range
    Dim ws As Worksheet
    Set ws = rng.Parent
    Set Rg_PickTilEnd = ws.range(rng.Address, rng.End(direction).Address)
End Function
Function A_MakeArrByRow(rngs As range) As Variant
'assuming that the for loop loop to the right
'Hard for chatGPT
    count = 0
    For Each cell In rngs
        If prev_row = cell.row Then
        
        Else
        
        End If
        If count = 0 Then
            Dim tempArr() As Variant
            tempArr = A_Append(cell.value)
        End If
        prev_row = cell.row
    Next
End Function
Sub S_TextJoinAt(rng_in, column_name As String)
    count = 0
    Dim outArr() As Variant
    For Each curr_cell In selection
        If count = 0 Then
        ' what if we choose only 1 model?
            Dim tempArr() As Variant
            tempArr = A_Append(tempArr, curr_cell.value)
        ElseIf count = selection.count - 1 Then
            If prev_row <> curr_cell.row Then
            'previous row
                fillAddr1 = column_name & prev_row
                Set toFillcurr_cell1 = range(fillAddr1)
                outStr1 = Join(tempArr, ", ")
                toFillcurr_cell1.value = outStr1
            'last row
                fillAddr2 = column_name & curr_cell.row
                Set toFillcurr_cell2 = range(fillAddr2)
                toFillcurr_cell2.value = curr_cell.value
            Else
                tempArr = A_Append(tempArr, curr_cell.value)
                fillAddr = column_name & curr_cell.row
                Set toFillcurr_cell = range(fillAddr)
                outStr = Join(tempArr, ", ")
                toFillcurr_cell.value = outStr
            End If
            
        Else
            
            If prev_row = curr_cell.row Then
            'Still the same row
                tempArr = A_Append(tempArr, curr_cell.value)
            Else
                'Next Row
                fillAddr = column_name & prev_row
                Set toFillcurr_cell = range(fillAddr)
                outStr = Join(tempArr, ", ")
                toFillcurr_cell.value = outStr
                tempArr = Array()
                tempArr = A_Append(tempArr, curr_cell.value)
            
            End If
        End If
        
        Set prev_cell = curr_cell
        prev_row = curr_cell.row
        count = count + 1
        
    Next
    
End Sub
Sub Bt_SelectModel()
'Hard for ChatGPT
    Set rng01 = selection
    Set leftCell = selection.Cells(1, 1)
    arr01 = A_toArray1d(selection)
    Call S_TextJoinAt(selection, "E")
    'arr02 = A_MakeArrByRow(Selection)
    'arr02 = A_toArray2d(Selection)
    'outStr = Join(rng01, ", ")
    'outStr = Join(arr01, ", ")
    'leftCell.Offset(0, -1) = outStr
    
End Sub

Sub Bt_AddMake()

End Sub

Sub Bt_ExportMake1()
'Done
    ws_temp01_name = "SR1.1"
    Set ws_template01 = Worksheets(ws_temp01_name)
    ws_make_name = "เลือกรถ"
    Set ws_make = Worksheets(ws_make_name)
    
    
    catch_Make_Arr = Array("Make", "ยี่ห้อ")
    catch_Model_Arr = Array("Model", "รุ่นรถ")
    Set text_make = Rg_FindAllRange(catch_Make_Arr, ws_temp01_name)
    If text_make.Offset(1, 0) <> "" Then
        Set old_makes = Rg_PickTilEnd(text_make.Offset(1, 0), xlDown)
        old_makes.ClearContents
    End If
    Set text_model = Rg_FindAllRange(catch_Model_Arr, ws_temp01_name)
    If text_model.Offset(1, 0) <> "" Then
        Set old_models = Rg_PickTilEnd(text_model.Offset(1, 0), xlDown)
        old_models.ClearContents
    End If
    Set text_make_ref = Rg_FindAllRange("ยี่ห้อ", ws_make_name)
    Set text_model_ref = Rg_FindAllRange("รุ่นรถที่เลือก", ws_make_name)
    
    make_list = A_FindFromHook("ยี่ห้อ", , , ws_make_name)
    model_list = A_FindFromHook("รุ่นรถที่เลือก", , , ws_make_name)
    Call A_FillValue(make_list, text_make.Offset(1, 0), , ws_temp01_name)
    Call A_FillValue(model_list, text_model.Offset(1, 0), , ws_temp01_name)
    MsgBox ("Export Succesfull !!")
End Sub

Sub Bt_ExportMake2()
'Not Done !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ws_temp02_name = "SR1.2"
    ws_make_name = "เลือกรถ"
    model_list = A_FindFromHook("รุ่นรถที่เลือก", , , ws_make_name)
    make_list = A_FindFromHook("ยี่ห้อ", , , ws_make_name)
    Dim model_arr() As Variant
    Dim count_arr() As Variant
    For i = LBound(model_list) To UBound(model_list)
        str01 = model_list(i)
        temp_arr = Split(str01, ", ")
        n = UBound(temp_arr) + 1
        count_arr = A_Append(count_arr, n)
        model_arr = A_Extend(model_arr, temp_arr)
        
    Next
    make_toFill = A_AddSpace(make_list, count_arr)
    If UBound(model_arr) > 30 Then
        MsgBox ("This template can't handle more than 30 models")
        Exit Sub
    End If
    Call S_UnmergedBetween2Cells("ยี่ห้อ", "หมายเหตุ", ws_temp02_name)
    
    catch_Make_Arr = Array("Make", "ยี่ห้อ")
    catch_Model_Arr = Array("Model", "รุ่นรถ")
    Set text_make = Rg_FindAllRange(catch_Make_Arr, ws_temp02_name)
    Set text_model = Rg_FindAllRange(catch_Model_Arr, ws_temp02_name)
    'Delete old content
    Set make_col = Rg_GetCellsBetween("ยี่ห้อ", "หมายเหตุ", ws_temp02_name)
    Set toDeleteRng = make_col.Resize(make_col.Rows.count, make_col.Columns.count + 1)
    toDeleteRng.ClearContents
    
    Call A_FillValue(make_toFill, text_make.Offset(1, 0), , ws_temp02_name)
    Call A_FillValue(model_arr, text_model.Offset(1, 0), , ws_temp02_name)
    Call Rg_Merged(text_make.Offset(1, 0), count_arr)
    Call W_BorderSubBulk(text_model.Offset(1, 0), count_arr)
    MsgBox ("Export Successful")

End Sub

Function A_toLongArray(rng, Optional delimiter = ",", Optional output_option = 0)
'Try to use delimiter to chop each string in the cells apart then put it into array
'Upgrade 01:  if I want to include many delimiters to detect?
'Upgrade 02: If rng is also Array?
    Dim outArr() As Variant
    For Each curr_cell In rng
        curr_str = curr_cell.value
        temp_arr = Split(curr_str, delimiter)
        For i = LBound(temp_arr) To UBound(temp_arr)
            temp_arr(i) = Trim(temp_arr(i))
        Next
        outArr = A_Extend(outArr, temp_arr)
    Next
    A_toLongArray = D_OutputFormat(outArr, output_option)
'How about try to do the oposite using number then output the input string ??
End Function
