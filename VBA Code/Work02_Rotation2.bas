Attribute VB_Name = "Work02_Rotation2"
Sub Test01()


    Dim ws As Worksheet
    Set ws = Worksheets("Luxury SUV G.2")

    Dim searchRange As range
    Set searchRange = ws.usedRange

    Dim text_cell As range
    str_pattern1 = "ความรับผิดต่อชีวิต"
    Set text_cell = searchRange.Find(str_pattern1)
    Set num_cell01 = W_FindUsedCellRight(text_cell)
    'Set row_used = F_UsedCellRow_Intersect(14)

    Set ans_cell = Sp_SelectFromTL(num_cell01, 3, 1)
    ans_arr = A_toArray1d(ans_cell)
    A_printArr (ans_arr)
    
End Sub
Sub test02()
    
End Sub

Function W_FindUsedCellRight(cell As range) As range
    Dim ws As Worksheet
    Set ws = cell.Worksheet

    Dim searchRange As range
    Set searchRange = ws.range(cell.Offset(0, 1), ws.Cells(cell.row, ws.Columns.Count).End(xlToLeft))

    Dim foundRange As range
    Set foundRange = searchRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)

    If Not foundRange Is Nothing Then
        Set W_FindUsedCellRight = foundRange
    Else
        Set W_FindUsedCellRight = Nothing
    End If
End Function



Function F_UsedCellRow(n As Long) As range
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or set ws = ThisWorkbook.Sheets("Sheet1")

    Dim searchRange As range
    Set searchRange = ws.range("A" & n & ":XFD" & n) ' search from A to XFD (last column)

    Dim foundRange As range
    Set foundRange = searchRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)

    If Not foundRange Is Nothing Then
        Set F_UsedCellRow = foundRange
    Else
        Set F_UsedCellRow = Nothing
    End If
End Function

Function F_UsedCellRow_Intersect(n As Long) As range
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or set ws = ThisWorkbook.Sheets("Sheet1")

    Dim rowRange As range
    Set rowRange = ws.range("A" & n & ":XFD" & n) ' range covering entire row

    Dim usedRange As range
    Set usedRange = ws.usedRange

    Dim intersectRange As range
    Set intersectRange = Intersect(rowRange, usedRange)

    If Not intersectRange Is Nothing Then
        Set F_UsedCellRow_Intersect = intersectRange.Cells(1, 1)
    Else
        Set F_UsedCellRow_Intersect = Nothing
    End If
End Function


