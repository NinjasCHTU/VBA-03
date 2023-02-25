Attribute VB_Name = "Scratch1"
Sub W_CheckBoxFromArray(arr As Variant)
    'Declare variables
    Dim frm As UserForm
    Dim cb As CheckBox
    Dim i As Long
    
    'Set the user form variable to the user form
    Set frm = UserForm1
    
    'Loop through the array
    For i = LBound(arr) To UBound(arr)
        'Add a check box to the user form
        Set cb = frm.CheckBoxes.Add(100, 100 + i * 25, 50, 20)
        'Set the caption of the check box to the element in the array
        cb.Caption = arr(i)
    Next i
    
End Sub

Sub WTest01()
    arr01 = Array(1, 2, 3, 4)
    Call W_CheckBoxFromArray(arr01)
    
End Sub
Sub W_Scratch01()
    UserForm1.Show
    'Declare a variable for the user form
    Dim frm As UserForm
    
    'Set the user form variable to the user form
    Set frm = UserForm1
    
    'Add a check box to the user form
    Set cb = frm.CheckBoxes.Add(100, 100, 50, 20)
End Sub
Sub O_SwapRanges()
    Set currSelect = selection
    
    Set range1 = currSelect.Areas(1)
    Set range2 = currSelect.Areas(2)
    n_row = range1.Rows.count
    n_col = range1.Columns.count
    Set holder = range("Q100:Z100")
    range1.Copy holder
    range2.Copy range1
    holder.Copy range2
    holder.Clear
End Sub

Sub Test05()
'GroupByColor
'assume that it's in the row position
    'Write Sub to Find total of Color
    'Copy and paste white color at far cell eg G50
    'Find the right most column with row 1
    'And do the same thing with other colors
    'Remember the original location
    'Then paste everything back to the original location

    
End Sub

Sub Test04()
    
    my2dArr1 = A_NumMatrix(3, 5)
    my2dArr2 = A_Reshape_2dTo1D(my2dArr1)
    A_printArr (my2dArr2)
    
   
End Sub



Sub Test03()
    arr01 = A_TxtTO1dArr("[1,2,3,4]")
    arr2 = A_Reshape_1dTo2D(arr01, 4, 1)
    MsgBox ("Done")
    
End Sub

Sub Test02()
    Dim arr03() As Variant
    arr01 = A_TxtTO1dArr("[1,2,3,4]")
    arr05 = A_TxtTO2dArr("[11,12,13],[21,22,23]")
    arr06 = A_TxtTO1dArr("[31,32,33]")
    arr05 = A_Append2(arr05, arr06)
    Call A_FillValue(arr05, range("K69"))
    'r01
    
End Sub



Sub test01()
    Call W_BorderOutside(range(G66))
    
End Sub







