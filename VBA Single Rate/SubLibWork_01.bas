Attribute VB_Name = "SubLibWork_01"
Sub W_CreateCheckBox()

    'Declare variables
    Dim ws As Worksheet
    Dim cb As CheckBox
    
    'Set the worksheet variable
    Set ws = ActiveSheet
    
    'Add a check box to the worksheet
    Set cb = ws.CheckBoxes.Add(100, 100, 50, 20, True, "")

End Sub

Sub W_CreateNameList()

    'Declare variables
    Dim ws As Worksheet
    Dim sel As range
    Dim lst As ListObject
    Dim i As Long
    
    'Set the worksheet and selection variables
    Set ws = ActiveSheet
    Set sel = Selection
    
    'Add a list object to the worksheet
    Set lst = ws.ListObjects.Add
    
    'Set the list object's data range to the selection, skipping the first row
    'Set lst.DataBodyRange = sel.Offset(1, 0).Resize(sel.Rows.Count - 1)
    
    'Clear the list object's table
    lst.DataBodyRange.Clear
    
    'Loop through the selection, skipping the first row
    For i = 2 To sel.Rows.count
        'If the value in the cell is not already in the list object, add it
        If Not lst.ListColumns(1).DataBodyRange.Find(sel.Cells(i, 1).value) Is Nothing Then
            lst.ListRows.Add
            lst.ListRows(lst.ListRows.count).range(1, 1).value = sel.Cells(i, 1).value
        End If
    Next i
    

End Sub
Sub W_DeleteGreyColumns()
      'Get the selected range
    Set sel = Selection
    
    'Get the used range
    Set used = ActiveSheet.usedRange
    
    'Find the intersection of the selected range and the used range
    Set row01 = Intersect(sel, used)
    
    Do Until Rg_IsGrey(row01) = False
      For Each cell In row01
        'Check if the cell's fill color is any shade of grey
        If Rg_IsGrey(cell) Then
          'If the cell's fill color is grey, delete the column
          cell.EntireColumn.Delete
          'Exit the loop and start over
          Exit For
        End If
      Next
    Loop
    
End Sub

