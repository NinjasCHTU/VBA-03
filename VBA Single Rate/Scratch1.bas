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
    Set currSelect = Selection
    
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
    path01 = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    Set wb01 = Wb_GetWB(path01)
    Set ws = wb01.Sheets(1)
    ws_name = ws.name
    arr01 = Array("เบี้ยสุทธิ")
    rng01 = Rg_FindAllRange(arr01)
    
    
    
End Sub

Sub Test02()
    defaultPath = "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct"
    Set wb01 = Wb_GetWB3(defaultPath)
    Set wb02 = Wb_GetWB3()
    
End Sub



Sub test01()
    Dim arr01_Up, arr01_Down, arr01_Left, arr01_Right, arr02 As range
    'arr01 = Pl_MotorCode("Standard1")
    'arr01 = Pl_Dealer_CoverageINFO("Standard1")
    'arr01 = Pl_Dealer_Premium("Standard1")
    arr01 = Pl_Insurer_CoverageINFO("Standard1")
    'Call A_FillValue(arr01, range("K70"), xlUp, xlToLeft)
    'Call A_FillValue(arr01, range("K70"), xlUp)
    'Call A_FillValue(arr01, range("K70"), xlDown, xlToLeft)
    'Call A_FillValue(arr01, range("K70"), xlDown)
    
    'Call A_FillValue(arr01, range("K70"), xlToLeft, xlUp)
    'Call A_FillValue(arr01, range("K70"), xlToLeft)
    'Call A_FillValue(arr01, range("K70"), xlToRight)
    Call A_FillValue(arr01, range("K70"), xlDown, , , True)
    
End Sub







