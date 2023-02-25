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
    n_row = range1.Rows.Count
    n_col = range1.Columns.Count
    Set holder = range("Q100:Z100")
    range1.Copy holder
    range2.Copy range1
    holder.Copy range2
    holder.Clear
End Sub

Sub Test005()
'GroupByColor
'assume that it's in the row position
    'Write Sub to Find total of Color
    'Copy and paste white color at far cell eg G50
    'Find the right most column with row 1
    'And do the same thing with other colors
    'Remember the original location
    'Then paste everything back to the original location

    
End Sub

Sub Test004()
    
    my2dArr1 = A_NumMatrix(3, 5)
    my2dArr2 = A_Reshape_2dTo1D(my2dArr1)
    A_printArr (my2dArr2)
    
   
End Sub



Sub Test003()
    arr1 = A_TxtTO1dArr("[1,2,3,4]")
    arr2 = A_ShiftRight(arr1, -2)
    A_printArr (arr2)
    
    'newArr1 = A_SetSubtract(myArr1, myArr2)
    'newArr2 = A_SetSubtract(myArr2, myArr1)
    'newArr3 = A_SetSubtract(myArr3, myArr5)
    'newArr4 = A_SetSubtract(myArr3, myArr4)
    
    'A_printArr (newArr1)
    'A_printArr (newArr2)
    'A_printArr (newArr3)
    'A_printArr (newArr4)
    
End Sub

Sub Test002()
    arr1 = A_TxtTO2dArr("[[1,1],[2,2],[3,3],[4,4]]")
    arr2 = A_TxtTO2dArr("[[a,aa],[b,bb],[c,cc],[d,dd]]")
    arr3 = A_HStack(arr1, arr2)
    temp = "Yeah almost Finish"
    
    
    
End Sub



Sub Test001()
    arr01 = Array(1, 2, 3, Array(4, 5, 6))



    Dim arr02(1 To 3, 1 To 4) As Double
    arr02(1, 1) = 1
    arr02(1, 2) = 2
    arr02(1, 3) = 3
    arr02(1, 4) = 4
    arr02(2, 1) = 5
    arr02(2, 2) = 6
    arr02(2, 3) = 7
    arr02(2, 4) = 8
    arr02(3, 1) = 9
    arr02(3, 2) = 10
    arr02(3, 3) = 11
    arr02(3, 4) = 12

    ans01 = A_isInArr2(arr01, 4)
    ans02 = A_isInArr2(arr02, 1)
    MsgBox (ans01)
    MsgBox (ans02)

    
    
    
    
End Sub







