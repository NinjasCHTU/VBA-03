Attribute VB_Name = "UnitTest01"
Sub test()
    newFileName = "Test01"
    Dim wb_NeoTemplate As Workbook
    Dim campaign_ws_name As String
    
    pathSheet = "ที่อยู่ไฟล์"
    
    folder = Rg_FindAllRange("Template Folder", pathSheet, xlPart).Offset(0, 1).value
    fileName = Rg_FindAllRange("Template File Name", pathSheet, xlPart).Offset(0, 1).value
    TemplatePath = folder & "\" & fileName
    outputFolder = Rg_FindAllRange("Output", pathSheet, xlPart).Offset(0, 1).value
    defaultFolder = Rg_FindAllRange("เลือกตาราง", pathSheet, xlPart).Offset(0, 1).value
    
    Set wb_NeoTemplate = Wb_GetWB3(, TemplatePath)
    Set wb_PremTable = Wb_GetWB3(defaultFolder)

    campaign_ws_name = wb_PremTable.Sheets(1).name
    'Open the file
    
    'outputFolder = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    
    
    
    '*************************************************************** Different ways to refer/open to wb_NeoTemplate  ***************************************************************'
    'Set wb_NeoTemplate = Wb_GetWB(defaultFolder)
    'Set wb_NeoTemplate = ThisWorkbook
    
    
    'Set wb_NeoTemplate = Workbooks("NeoTemplate 04_ForVBA.xlsx")
    
    'Set wb_NeoTemplate = Workbooks.Open()
    '*************************************************************** Different ways to refer/open to wb_NeoTemplate  ***************************************************************'
    
    'Save the file as a new file
    Dim newTemplatePath As String
    newTemplatePath = outputFolder & "\" & newFileName
    newTemplatePath2 = newTemplatePath & ".xlsx"
    If Dir(newTemplatePath2) = "" Then
'What if the file is already opened?
        'MsgBox "File not found."
    Else
'        On Error GoTo Err01
'        Set old_file = Workbooks(newFileName)
'        old_file.Close False
'        On Error GoTo 0
'        On Error GoTo -1
Err01:
        Kill newTemplatePath2
    End If
    ' Don't show confirmation window
    Application.DisplayAlerts = False
    wb_NeoTemplate.SaveAs fileName:=newTemplatePath
    ' Allow confirmation windows to appear as normal
    Application.DisplayAlerts = True
    wb_NeoTemplate.Close
End Sub
Sub test_Ws_WS_at_WB()
    
    ws_str = "SR1.1"
    Set wb01 = ThisWorkbook
    wb_str = wb01.name
    Set ws01 = Sheets(ws_str)
    Set ans01 = Ws_WS_at_WB(ws_str, wb01)
    ans01_02 = Ws_WS_at_WB(ws_str, wb01, False)
    Set ans02 = Ws_WS_at_WB(ws_str, wb_str)
    Set ans03 = Ws_WS_at_WB(ws01, wb01)
    Set ans04 = Ws_WS_at_WB(ws01, wb_str)
End Sub
Sub test_Rg_FindAllRanges()
    df_path = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    Set wb01 = Wb_GetWB3(df_path)
    
    Set ws = wb01.Sheets(1)
    ws_name = ws.name
    wb01.Activate
    Set rng01 = Rg_FindAllRange("เบี้ยสุทธิ", ws_name, xlPart, wb01)
    
    
End Sub
Sub test_A_HStack()
    arr01 = Array(1, 3)
    arr03 = A_NumMatrix(2, 4)
    Dim arr02() As Variant

    ans01 = A_HStack(arr01, arr01)
        '1d to 1d
    ans02 = A_HStack(arr03, arr02)
    '2d with empty
    ans03 = A_HStack(arr02, arr03)
    '2d with empty swap position
    ans04 = A_HStack(arr03, arr01)
    ans05 = A_HStack(arr01, arr03)

End Sub
Sub test_A_GetRow()
    num01 = A_TxtTO2dArr("[ [1,2], [3,4]")
    Dim x As Long
    x = 0
    ans01 = A_GetRow(num01, x)
    ans02 = A_GetRow(num01, 1)
End Sub
Sub test_A_GetColumn()
    num01 = A_TxtTO2dArr("[ [1,2], [3,4]")
    Dim x As Long
    x = 0
    ans01 = A_GetColumn(num01, x)
    ans02 = A_GetColumn(num01, 1)
End Sub
Sub test_A_Replicates()
    num01 = A_TxtTO2dArr("[ [1,2], [3,4]")
    num02 = A_TxtTO2dArr("[ [1,2], [3,4],[5,6]")
    'ans01_01 = A_Replicates(num01, 3)
    'ans01_02 = A_Replicates(num01, 3, xlToRight)
    'ans02_01 = A_Replicates(num01, 3, xlDown, True, True)
    'ans03_01 = A_Replicates(num01, 3, xlToRight, True, True)
    'ans02_02 = A_Replicates(num01, 3, xlDown, False, True)
    ans03_02 = A_Replicates(num02, 3, xlToRight, False, True)
    ws_name = "Debug3"
    Set rng01 = Sheets("Debug3").range("AE2")
    'A_printArr (ans01)
    'Call A_FillValue(ans01, rng01, , , True, ws_name)
    'Call A_FillValue(myArray:=ans01_01, start_cell:=rng01, ws_name:=ws_name, overwrite:=True)
    'Call A_FillValue(myArray:=ans01_02, start_cell:=rng01, ws_name:=ws_name, overwrite:=True)
    'Call A_FillValue(myArray:=ans02_01, start_cell:=rng01, ws_name:=ws_name)
    'Call A_FillValue(myArray:=ans02_02, start_cell:=rng01, ws_name:=ws_name)
   'Call A_FillValue(myArray:=ans03_01, start_cell:=rng01, ws_name:=ws_name)
    Call A_FillValue(myArray:=ans03_02, start_cell:=rng01, ws_name:=ws_name)
End Sub
Sub test_A_IsEmpty()
    Dim arr1D() As Variant
    ReDim arr1D(4)
    Dim arr1D_2() As Variant
    
    arr1DNonEmpty = Array("apple", "banana", "orange")
    Dim arr2D() As Variant
    ReDim arr2D(0 To 3, 0 To 2) As Variant
    arr2DNonEmpty = A_TxtTO2dArr(" [red,blue],[green,yellow] ")
    
    ans1Empty = A_IsEmpty2(arr1D)
    ans1Empty_2 = A_IsEmpty2(arr1D_2)
    ans2NonEmpty = A_IsEmpty2(arr1DNonEmpty)
    ans3_2D = A_IsEmpty2(arr2D)
    arr4_2DNonEmpty = A_IsEmpty2(arr2DNonEmpty)
End Sub
Sub test_Rg_RangeFromText()
    ws_name = "SR1.1"
    Set rng01 = Rg_RangeFromText("จำนวนผู้ขับขี่", "รหัสรถ 210", ws_name)
    Set rng02 = Rg_RangeFromText("จำนวนผู้ขับขี่", "รหัสรถ 210", ws_name, xlWhole)
End Sub
Sub test_Rg_RangeFromRange()
    ws_name = "SR1.1"
    Set rng01 = Sheets(ws_name).range("B14")
    Set rng02 = Sheets(ws_name).range("H12")
    Set ans01 = Rg_RangeFromRange(rng01, rng02)
    Set ans02 = Rg_RangeFromRange(rng02, rng01)
    
End Sub
Sub test_A_Extend()
    arr01 = A_NumMatrix(3, 4)
    arr02 = A_NumMatrix(5, 4)
    arr03 = Array(1, 2, 3, 4)
    arr04 = Array("a", "b", "c", "d")
    Set rng01 = range("E21")
    
    ws_name = "Next&Button"
    ans01 = A_Extend(arr1, arr02)
    Call A_FillValue(ans01, rng01, , , ws_name, True)
    ans02 = A_Append2(ans01, arr03)
    Call A_FillValue(ans02, rng01, , , ws_name, True)
    ans03 = A_Append2(ans02, arr04)
    Call A_FillValue(ans03, rng01, , , ws_name, True)
End Sub
Sub test_A_ShiftIndexTo0()
    arr01 = A_NumMatrix(3, 4)
    arr02 = VB_Transpose(arr01)
    ans01 = A_ShiftIndexTo0(arr02)
    ws_name = "Next&Button"
    Call A_FillValue(ans01, range("E21"), , , ws_name, True)
    
End Sub
Sub test_W_GetNum1Line()
    ws_name = "SR1.1"
    Set rng01 = Sheets(ws_name).range("H23")
    ans01 = W_GetNum1Line(rng01)
End Sub

Sub test_A_FillValue()
    arr01 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Set rng01 = range("E21")
    Set rng02 = range("K21")
    ws_name = "Next&Button"
    'Call A_FillValue(arr01, rng01, xlDown, , ws_name)
    'Call A_FillValue(arr01, rng01, xlToRight, , ws_name, True)
    'Call A_FillValue(arr01, rng02, xlToLeft, , ws_name, True)
    'Call A_FillValue(arr01, rng01, xlUp, , ws_name)
    '************************************************************************** Test for 2d Array *********************************************************************************
    arr02 = A_NumMatrix(3, 4)
    Call A_FillValue(arr02, rng01, , , ws_name)
    Call A_FillValue(arr02, rng01, , True, ws_name)
    
End Sub
Sub test_Rg_Resize()
    Dim rng01, rng02 As range
    Set rng01 = range("E21")
    Set rng02 = range("E21:O28")
    
    Set ans01 = Rg_Resize(rng01, 2, 3)
    Set ans02 = Rg_Resize(rng01, 2, -3)
    Set ans03 = Rg_Resize(rng01, -2, 3)
    Set ans04 = Rg_Resize(rng01, -2, -3)
    
    Set ans01 = Rg_Resize(rng02, 2, 3)
    Set ans02 = Rg_Resize(rng02, 2, -3)
    Set ans03 = Rg_Resize(rng02, -2, 3)
    Set ans04 = Rg_Resize(rng02, -2, -3)
End Sub
Sub test_Rg_NextContainNum()
    Dim rng01, rng02 As range
    Set ws01 = Sheets("SR1.1")
    Set rng01 = ws01.range("H19")
    Set rng02 = ws01.range("K14")
    Set rng03 = ws01.range("K50")
    
    Set ans01 = Rg_NextContainNum(rng01, xlToRight)
    Set ans02 = Rg_NextContainNum(rng01, xlToLeft)
    Set ans03 = Rg_NextContainNum(rng01, xlUp)
    Set ans04 = Rg_NextContainNum(rng01, xlDown)
    
    
End Sub
Sub test_Rg_NextNumeric()
    Dim rng01, rng02 As range
    Set ws01 = Sheets("SR1.1")
    Set rng01 = ws01.range("H18")
    Set rng02 = ws01.range("K14")
    Set rng03 = ws01.range("K50")
    
    Set ans01 = Rg_NextNumeric(rng01, xlToRight)
    Set ans02 = Rg_NextNumeric(rng01, xlToLeft)
    Set ans03 = Rg_NextNumeric(rng01, xlUp)
    Set ans04 = Rg_NextNumeric(rng01, xlDown)
    
    
End Sub
Sub test_Rg_NextTextCell()
    Dim rng01, rng02 As range
    Set ws01 = Sheets("SR1.1")
    Set rng01 = ws01.range("J14")
    Set rng02 = ws01.range("K14")
    Set rng03 = ws01.range("K50")
    
    Set ans01 = Rg_NextTextCell(rng01, xlToRight)
    Set ans02 = Rg_NextTextCell(rng02, xlToRight)
    Set ans03 = Rg_NextTextCell(rng03, xlDown)
End Sub

Sub test_Wb_GetWB()
    defaultPath = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    Set wb01 = Wb_GetWB(defaultPath)
    Set wb02 = Wb_GetWB(defaultPath)
End Sub
Sub test_Rg_FindSomeRanges()
'Not Done debugging
    df_path = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    Set wb01 = Wb_GetWB3(df_path)
    
    Set ws = wb01.Sheets(1)
    ws_name = ws.name
    wb01.Activate

    arr02 = Array(-1)
    Set rng01 = Rg_FindSomeRanges("รหัส", -1, ws_name, , , wb)
    
    Set rng02 = Rg_FindSomeRanges("รหัส", Array(-1), ws_name, , , wb)
    Set rng03 = Rg_FindSomeRanges("รหัส", Array(1, 2, 4, -2), ws_name, , , wb)
    Set rng04 = Rg_FindSomeRanges("รหัส", Array(-1, -3), ws_name, , , wb)
    Set rng05 = Rg_FindSomeRanges("รหัส", 3, ws_name, , , wb)
    
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

