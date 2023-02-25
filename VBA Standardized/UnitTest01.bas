Attribute VB_Name = "UnitTest01"
Function UnionInFunction()
'This is just a test function not a real function
'To show that Union actually work in function
    Set temp01 = ThisWorkbook.Sheets("Sheet2").range("A15:A100")
    Set temp02 = ThisWorkbook.Sheets("Sheet2").range("C15:C100")
    Set outRng = Union(temp01, temp02)
    Set UnionInFunction = outRng
End Function
Sub test_UnionInFunction()
    
    Set var01 = UnionInFunction()
End Sub
Sub test_Bt_UpdateRedbook()
'Not Done just started !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    Dim filepath As Variant
    Dim wb_redbook As Workbook
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    On Error GoTo Err01
    Set xlApp = New Excel.Application
    On Error GoTo Err02
    'ChDir "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    ChDir "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    On Error GoTo 0
    'Change default path when open the select file window
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    'FileFilter = make the user sees only excel file
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    Set xlApp = New Excel.Application
    If filepath <> False Then
    'UpdateLinks:=0 not to update links for redbookfile
        Set xlWorkbook = xlApp.Workbooks.Open(filepath, UpdateLinks:=0)
    Else
        MsgBox ("The import is canceled")
        Exit Sub
    End If
    Application.ScreenUpdating = True
    On Error GoTo Err01
    ws01 = Wb_byUsedRange(xlWorkbook)
    ' ws_redbook has the most used cells(with text)
    Set ws_redbook = ws01(0)
    Call test_W_GetRedBookData(ws_redbook)
    
    MsgBox ("Make Model Import Succesful !!!")
    xlWorkbook.Close
    MsgBox ("File is closed")
    On Error GoTo 0
    Exit Sub
Err01:
    MsgBox ("There's an error during make model import")
    xlWorkbook.Close
    MsgBox ("File is closed")
    Exit Sub
Err02:
    MsgBox ("There's a problem with default path check the VBA code at Bt_UpdateRedbook")
    Exit Sub
End Sub
Sub test_W_GetRedBookData(ws)
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
    'Set col_make = Rg_PickTilEnd(start_make, xlDown)
    'Set col_model = Rg_PickTilEnd(start_model, xlDown)
    'Set col_group = Rg_PickTilEnd(start_group, xlDown)
    Set temp01 = ThisWorkbook.Sheets("Sheet2").range("A15:A100")
    Set temp02 = ThisWorkbook.Sheets("Sheet2").range("A15:A100")
    Set combined01 = Union(temp01, temp02)
    
    Set col_make = ws.range("B2:B100")
    Set temp = ws.range("B101:B200")
    'Set col_model = ws.range("E2:E100")
    'Set col_group = ws.range("H2:H100")
    'Set col_make = ws.Columns(col_make_name)
    'Set col_model = ws.Columns(col_model_name)
    'Set col_group = ws.Columns(col_group_name)
    Set combinedCol = Union(col_make, temp)
    'Set combinedCol = Rg_Union(col_make, temp)
    'Set combinedCol = Union(col_make, col_model, col_group)
    outArr = D_UniqueRow(col_make, col_model, col_group)
    arr_make = col_make.value
    arr_model = col_model.value
    arr_group = col_group.value
End Sub
Sub test_W_BorderSubBulk()
    Set start01 = Sheets("Sheet2").range("G21")
    Set start02 = Sheets("Sheet2").range("H21")
    num01 = Array(3, 2, 4, 2)
    num02 = Array(1, 1, 1, 2, 3, 4, 3, 2, 1, 1, 1, 3, 1)
    Call W_BorderSubBulk(start01, num01)
    Call W_BorderSubBulk(start02, num02)
End Sub
Sub test_A_FillValue()
    Dim arr01() As Variant
    For i = 1 To 40
        arr01 = A_Append(arr01, i)
    Next
    Set rng01 = Sheets("SR1.2").range("C83")
    Call A_FillValue(arr01, rng01, xlUp)
    
End Sub
Sub test_W_BorderSubElement()
    Call W_BorderSubElement(selection)
End Sub
Sub test_W_BorderOutside()
    Call W_BorderOutside(selection)
End Sub
Sub test_A_AddSpace()
    arr01 = Array("a", "b", "c", "d")
    num01 = Array(1, 2, 3, 4)
    ans01 = A_AddSpace(arr01, num01)
    ans02 = A_AddSpace(arr01, num01, 1)
End Sub
Sub test_Rg_Merged()
    arr01 = Array(1, 2, 3, 4)
    arr02 = Array(3, 2, 4, 2)
    arr03 = Array(2, 1, 1, 1, 3, 4, 2)
    Set rng01 = Sheets("Sheet2").range("G21")
    Set rng02 = Sheets("Sheet2").range("H21")
    Set rng03 = Sheets("Sheet2").range("I21")
    Call Rg_Merged(rng01, arr01)
    Call Rg_Merged(rng02, arr02)
    Call Rg_Merged(rng03, arr03)
End Sub
Sub test_D_AlphaBigHelp01()
    ans01 = D_AlphaBigHelp01(4)
    ans02 = D_AlphaBigHelp01(10)
    ans03 = D_AlphaBigHelp01(100)
    ans04 = D_AlphaBigHelp01(3600)
End Sub
Sub test_Rg_LeftOf()
'test_Rg_PickTopOf
'test_Rg_PickBottomOf
'test_Rg_PickRightOf
    Dim rng01, rng02 As range
    Set rng01 = range("A1:E20")
    Set rng02 = range("C10")
    str01 = "A1:E20"
    str02 = "C10"
    Set ans01_01 = Rg_PickTopOf(rng01, rng02)
    Set ans01_02 = Rg_PickTopOf(str01, str02)
    
    Set ans02_01 = Rg_PickBottomOf(rng01, rng02)
    Set ans02_02 = Rg_PickBottomOf(str01, str02)
    
    Set ans03_01 = Rg_PickLeftOf(rng01, rng02, 0)
    Set ans03_02 = Rg_PickLeftOf(str01, str02)
    
    Set ans03_01 = Rg_PickRightOf(rng01, rng02)
    Set ans03_02 = Rg_PickRightOf(str01, str02)
    
End Sub
Sub test_Rg_LeftORRight()
    Set rng01 = range("A1")
    Set rng02 = range("B20")
    Set rng03 = range("B2")
    Set rng04 = range("E3")
    Set rng05 = range("O3")
    
    ans01 = Rg_LeftORRight(rng01, rng02)
    ans02 = Rg_LeftORRight(rng02, rng03)
    ans03 = Rg_LeftORRight(rng04, rng03)
    
    ans04 = Rg_TopORBottom(rng01, rng02)
    ans05 = Rg_TopORBottom(rng02, rng01)
    ans06 = Rg_TopORBottom(rng04, rng05)
    
End Sub
Sub test_Dc_toDict()
'Already work with both range and array
    Dim key_array As Variant
    key_array = Array("key1", "key2", "key3", "key4")
    
    Set rng01 = range("F26:F29")
    Set rng02 = range("G26:G29")
    Dim val_array As Variant
    val_array = Array("value1", "value2", "value3", "value4")
    
    Dim dict As Object
    Set dict = Dc_toDict(key_array, val_array)
    Set dict02 = Dc_toDict(rng01, rng02)
    
    For Each key In dict.Keys
      Debug.Print key, dict(key)
    Next key
    For Each key In dict02.Keys
      Debug.Print key, dict(key)
    Next key

End Sub

Sub test_A_FindFromHook()
'THAI FIX
    arr01 = Array("ยี่ห้อ", "make")
    arr02 = Array("ยี่ห้อ", "รุ่นรถ")
    arr03 = Array("เบี้ย")
    
    str01 = "make"
    str02 = "kenfe"
    str03 = "เบี้ย"
    ws01 = "SingleRate2"
    ans01 = A_FindFromHook(arr01, , , ws01)
    ans02 = A_FindFromHook(str01, , , ws01)
    ans03 = A_FindFromHook(str02, , , ws01)
    
    ans04 = A_FindFromHook(arr02, , , ws01)
    ans05 = A_FindFromHook(arr03, , , ws01)
    ans06 = A_FindFromHook(str03, , , ws01)
    MsgBox ("test_A_FindFromHook Done")

End Sub
Sub test_St_ContainsNum()
    Dim str01, str02, str03 As String
    str01 = "ienfno ef"
    str02 = "206446 "
    str03 = "dfefe 6164"
    str04 = "2154545 บาท/คน"
    ans01 = St_ContainsNum(str01)
    ans02 = St_ContainsNum(str02)
    ans03 = St_ContainsNum(str03)
    ans04 = St_ContainsNum(str04)
    MsgBox ("Done")
End Sub
Sub test_St_GetBefore()
    Dim text01, text02 As String
    text01 = "6384692 io free"
    text02 = "93489_ jekf"
    arr01 = Array(" ", "_")
    
    ans01 = St_GetBefore(text01, arr01)
    ans02 = St_GetBefore(text01, " ")
    ans03 = St_GetBefore(text01, "_")
    MsgBox (ans01)
    MsgBox (ans02)
    MsgBox (ans03)
    'MsgBox (ans01)
End Sub
Sub test_A_Count()
    arr01 = Array(1, 1, 1, 2, 2)
    arr02 = A_NumMatrix(3, 4)
    
    element01 = Array(11, 13, 52)
    element02 = A_NumMatrix(2, 2)
    
    
    ans01 = A_Count(arr01, 2)
    ans02 = A_Count(arr01)
    ans03 = A_Count(arr02, element01)
    ans04 = A_Count(arr02, element02)
    'MsgBox (ans01)
    'MsgBox (ans02)
    'MsgBox (ans03)
    MsgBox (ans04)
    
End Sub

Sub test_A_Append2()
    Dim arr03() As Variant
    arr01 = Array(1, 1, 2, 2)
    arr02 = A_NumMatrix(3, 4)
    
    element01 = Array(11, 13, 52)
    element02 = A_NumMatrix(2, 2)
    
    ans01 = A_Append2(arr02, arr01)
    ans02_01 = A_Append2(arr01, 700)
    ans02_02 = A_Append2(arr03, 50)
    ans03_01 = A_Append2(arr03, 4)
    ans03_02 = A_Append2(arr03, arr01)
    
    MsgBox ("Done")

    
End Sub

