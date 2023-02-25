Attribute VB_Name = "SubLib_01"
Sub ColorGroups()
''' Not Done (And not that Important)
''' Difficult for ChatGPT
    Dim cell As range
    Dim color_arr As Variant
    Dim colorIndex As Long
    color_arr = Array( _
        "FFCCFF", _
        "FFCCCC", _
        "FFCC99", _
        "FFCC66", _
        "FFCC33", _
        "FFCC00", _
        "FFFFCC", _
        "FFFF99", _
        "FFFF66", _
        "FFFF33", _
        "FFFF00", _
        "CCFFFF", _
        "CCCCFF", _
        "CC99FF", _
        "CC66FF", _
        "CC33FF", _
        "CC00FF", _
        "99CCFF", _
        "9966FF", _
        "3333FF")
    Dim UniqueWords() As Variant
    
    Set word_count = D_DictCount(selection)
    For Each cell In selection
        
        If cell.value <> "" Then
            curr_word = cell.value
            
            If A_isInArr(UniqueWords, curr_word) Then
                color_group = A_FindIndex(UniqueWords, curr_word)
                cell.Interior.color = CLng("&H" & color_arr(color_group))
            Else
                UniqueWords = A_Append(UniqueWords, curr_word)
            End If
        End If
    Next cell
End Sub

Sub S_toValue()
    For Each cell In selection
        cell.value = cell.value
    Next cell
End Sub
Sub S_NumToText()
    For Each cell In selection
        cell.NumberFormat = "@"
    Next cell
End Sub

Sub S_RotateFontColor()
    Set myRange = selection
    color_arr1 = Rg_ExtractFontColor(myRange)
    color_arr2 = A_ShiftRight(color_arr1, -1)
'Rg_ChangeFillColors changes fill color
'But it works here... worth investigating more
    Call Rg_ChangeFontColors(myRange, color_arr1, color_arr2)
    
End Sub
Sub S_RotateFillColor()
    Set myRange = selection
    color_arr1 = Rg_ExtractFillColor(myRange)
    color_arr2 = A_ShiftRight(color_arr1, -1)
'Rg_ChangeFillColors changes fill color
'But it works here... worth investigating more
    Call Rg_ChangeFillColors(myRange, color_arr1, color_arr2)
    
End Sub

Sub S_DragFormulaDown()
' This task is difficult for GPT !!!!!!!!!!!!!!!!!!!!!!
    Set range1 = selection.Areas(1)
    Set range2 = selection.Areas(2)
    
    If range1.count <= range2.count Then
        Set range_formula = range1
        Set range_runDown = range2
    Else
        Set range_formula = range2
        Set range_runDown = range1
    End If
    
    
    i = 1
    start_formula = range_formula.Formula
    start_cell = Replace(range_runDown.Cells(1, 1).Address, "$", "")
    For Each cell In range_runDown
        If i > 1 Then
            curr_addr = Replace(cell.Address, "$", "")
            myFormula = Replace(start_formula, start_cell, curr_addr)
            range_formula.Offset(i - 1, 0).Formula = myFormula
        End If
        i = i + 1
    Next cell
    

End Sub

Sub S_DragFormulaRight()
' This task is difficult for GPT !!!!!!!!!!!!!!!!!!!!!!
    Set range1 = selection.Areas(1)
    Set range2 = selection.Areas(2)
    
    If range1.count <= range2.count Then
        Set range_formula = range1
        Set range_runDown = range2
    Else
        Set range_formula = range2
        Set range_runDown = range1
    End If
    
    
    i = 1
    start_formula = range_formula.Formula
    start_cell = Replace(range_runDown.Cells(1, 1).Address, "$", "")
    For Each cell In range_runDown
        If i > 1 Then
            curr_addr = Replace(cell.Address, "$", "")
            myFormula = Replace(start_formula, start_cell, curr_addr)
            range_formula.Offset(0, i - 1).Formula = myFormula
        End If
        i = i + 1
    Next cell
    

End Sub
Sub S_CopyRightIfColorFont()
    Set range_in = selection
    For Each celnl In range_in
        If cell.Font.colorIndex <> 1 Then
            cell.Copy Destination:=cell.Offset(0, 1)
        End If
    Next cell
End Sub
Sub S_CopyRightIfFill()
    Set range_in = selection
    For Each cell In range_in
        If cell.Interior.colorIndex <> -4142 And cell.Interior.colorIndex <> 7 Then
            cell.Copy Destination:=cell.Offset(0, 1)
        End If
    Next cell
End Sub
Sub S_DeleteNoColorCell()
    Set range_in = selection
    
    
    For Each cell In range_in
        If cell.Interior.colorIndex = -4142 Then
            cell.Clear
        End If
    Next cell
    
End Sub

Sub S_DeleteColorCell()
    Set range_in = selection
    
    
    For Each cell In range_in
        If cell.Interior.colorIndex <> -4142 Then
            cell.Clear
        End If
    Next cell
    
End Sub


Sub S_DeleteNoColorCellShift()
    Set range_in = selection
    
    Do While Rg_HasUnfilledCells(range_in)
        For Each cell In range_in
            If cell.Interior.colorIndex = -4142 Then
                cell.Delete shift:=xlShiftUp
            End If
        Next cell
    Loop
    

End Sub

Sub S_DeleteColorCellShift()
' Have problems...
    Set range_in = selection
    
    Do While Rg_HasColorCells(range_in)
        For Each cell In range_in
            If cell.Interior.colorIndex <> -4142 Then
                cell.Delete shift:=xlShiftUp
            End If
        Next cell
    Loop
    Call S_DeleteShiftUp
    

End Sub


Sub S_DeleteShiftUp()
    Set range_in = selection
    
    Do While WorksheetFunction.CountBlank(range_in) > 0
        For Each cell In range_in
            If cell.value = "" Then
                cell.Delete shift:=xlShiftUp
            End If
        Next cell
    Loop
    

End Sub
Sub S_ColorFontElement()
'Chat GPT
'selection 1 is the range with color
'selection 2 is the cell that I want to change fill color
    Set ColorArea = selection.Areas(1)
    Set BlankArea = selection.Areas(2)
    For Each cell_color In ColorArea
        For Each cell_blank In BlankArea
            If cell_color.value = cell_blank.value Then
                cell_blank.Font.color = cell_color.Font.color
            End If
        Next cell_blank
    Next cell_color


End Sub
Sub S_ColorElement()
'Chat GPT
'selection 1 is the range with color
'selection 2 is the cell that I want to change fill color
    Set ColorArea = selection.Areas(1)
    Set BlankArea = selection.Areas(2)
    For Each cell_color In ColorArea
        For Each cell_blank In BlankArea
            If cell_color.value = cell_blank.value Then
                cell_blank.Interior.color = cell_color.Interior.color
            End If
        Next cell_blank
    Next cell_color


End Sub
Sub S_CopyExactFormula()
    Set rngSource = selection.Areas(1)
    Set rngDestination = selection.Areas(2)
    Call S_CopyExactFormula_Help01(rngSource, rngDestination)
    

End Sub
Sub S_CopyExactFormula_Help01(rngSource, rngDestination)
''''''''''''''' FIX  Have a problem when the destination many cells
'''''''''''' for now it works when selection only 1 cell in destination

    Dim i As Long
    Dim n_row As Long
    
    
    n_row = rngSource.Rows.count
    n_col = rngSource.Columns.count
    
    Set topLeft_Destination = Sp_SelectFromTL(rngDestination)
    For i = 0 To n_row - 1
        For j = 0 To n_col - 1
            
            myFormula = rngSource.Cells(i + 1, j + 1).FormulaArray

            
            rngDestination.Offset(i, j).Formula = myFormula
        Next j
    Next i
    
End Sub

Sub S_ChangeFontColor()
' Change FromColor to ToColor
' Note Selection order: myArea,Color1,Color2
    Set myArea = selection.Areas(1)
    Set FromColor = selection.Areas(2)
    Set ToColor = selection.Areas(3)
    Call S_ChangeFontColorHelp1(myArea, FromColor, ToColor)
    
End Sub


Sub S_ChangeFontColorHelp1(myArea, FromColor, ToColor)
    For Each cell In myArea
        If cell.Font.color = FromColor.Font.color Then
            cell.Font.color = ToColor.Font.color
        End If
    Next cell
End Sub


Sub S_ChangeFillColor()
' Change FromColor to ToColor
' Note Selection order: myArea,Color1,Color2
    Set myArea = selection.Areas(1)
    Set FromColor = selection.Areas(2)
    Set ToColor = selection.Areas(3)
    Call S_ChangeFillColorHelp1(myArea, FromColor, ToColor)
    
End Sub


Sub S_ChangeFillColorHelp1(myArea, FromColor, ToColor)
    For Each cell In myArea
        If cell.Interior.color = FromColor.Interior.color Then
            cell.Interior.color = ToColor.Interior.color
        End If
    Next cell
End Sub


Sub S_SwapFontColor()
' Swap color1 to color2 and vice versa in the area called myArea
' Note Selection order: myArea,Color1,Color2
    Set myArea = selection.Areas(1)
    Set Color1 = selection.Areas(2)
    Set Color2 = selection.Areas(3)
    Call S_SwapFontColorHelp1(myArea, Color1, Color2)
    
End Sub

Sub S_SwapFontColorHelp1(myArea, Color1, Color2)
    For Each cell In myArea
        If cell.Font.color = Color1.Font.color Then
            cell.Font.color = Color2.Font.color
        ElseIf cell.Font.color = Color2.Font.color Then
            cell.Font.color = Color1.Font.color
        End If
    Next cell
End Sub

Sub S_SwapFillColor()
' Swap color1 to color2 and vice versa in the area called myArea
' Note Selection order: myArea,Color1,Color2
    Set myArea = selection.Areas(1)
    Set Color1 = selection.Areas(2)
    Set Color2 = selection.Areas(3)
    Call S_SwapFillColorHelp1(myArea, Color1, Color2)
    
End Sub

Sub S_SwapFillColorHelp1(myArea, Color1, Color2)
    For Each cell In myArea
        If cell.Interior.color = Color1.Interior.color Then
            cell.Interior.color = Color2.Interior.color
        ElseIf cell.Interior.color = Color2.Interior.color Then
            cell.Interior.color = Color1.Interior.color
        End If
    Next cell
End Sub

Sub MakeMyLibTemplateFile()
'From ChatGPT !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ' Temporarily disable delete confirmation dialog
    Application.DisplayAlerts = False
    
    ' Delete all the sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Delete
    Next ws
    
    ' Add a new sheet named "sheet1"
    ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    ActiveSheet.name = "sheet1"
    
    ' Restore delete confirmation dialog
    Application.DisplayAlerts = True


End Sub
Sub S_deleteEmptyRows()
    Dim i As Integer
    Dim intLastRow As Integer
    
    intLastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).row
    For i = intLastRow To 1 Step -1
        If Application.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
        End If
    Next
    
End Sub

Sub S_SuperscriptBulk(inRange As range, inArr As Variant)
    n = UBound(inArr)
    For i = 0 To n
        curr_inx = inArr(i)
        With inRange.Characters(start:=curr_inx, Length:=1).Font
            .Superscript = True
        End With
        
    Next i
End Sub
Sub S_SubscriptBulk(inRange As range, inArr As Variant)
    n = UBound(inArr)
    For i = 0 To n
        curr_inx = inArr(i)
        With inRange.Characters(start:=curr_inx, Length:=1).Font
            .Subscript = True
            
        End With
        
    Next i
End Sub


Sub PrintColorCode()

End Sub
Sub S_ColorIf()
'Can run but need more features
    Dim arr_of_checker() As Variant
    arr_of_checker = Array("a", "e", "i", "o", "u", "h")
    Dim inRange As range
    Set inRange = selection
    For Each curr_cell In inRange
        
        curr_string = curr_cell.value
        first_ch = Left(curr_string, 1)
        first_ch = St_UnDiaCriticVB(first_ch)
        If A_isInArr(arr_of_checker, first_ch) Then
            curr_cell.Interior.color = vbYellow
        End If
        
    Next curr_cell
    
    
End Sub
Sub S_ColorFontFromTo()

End Sub
Sub S_ColorFontAt()
    Dim inRange As range
    On Error Resume Next
    Set inRange = Application.InputBox(Prompt:="Please select your range for coloring", Type:=8)
    On Error GoTo 0
    If inRange Is Nothing Then Exit Sub
    
    
    inx = InputBox("Enter the index: ")
    'inx = 3
    For Each curr_cell In inRange
        With curr_cell.Characters(start:=inx, Length:=1).Font
            .color = vbBlue
        End With
    Next curr_cell
    
    

End Sub
'Global wordToColor As Variant
'Global n_word As Integer
'Global word_list As Variant
Sub S_ColorString1()
    Dim sentence_list As range
    On Error Resume Next
    Set sentence_list = Application.InputBox(Prompt:="Please select your range for coloring", Type:=8)
    Set word_list = Application.InputBox(Prompt:="Please select COLERED WORD", Type:=8)
    On Error GoTo 0
    If sentence_list Is Nothing Then Exit Sub
    If word_list Is Nothing Then Exit Sub
    
    singleWord = word_list.value
    n_word = Len(singleWord)
    myColor = word_list.Font.color
    Call S_ColorStringTask(singleWord, n_word, myColor, sentence_list)
End Sub

Sub S_ColorStringTask(ByRef singleWord, ByRef n_word, ByRef myColor, ByRef sentence_list As range)
'Add: More words
'Add: Custom color
'Add: Color multipleTimes
    
    
    For Each curr_cell In sentence_list
        curr_str = curr_cell.value
        inx = InStr(1, curr_str, singleWord, vbTextCompare)
        If inx <> 0 Then
            With curr_cell.Characters(start:=inx, Length:=n_word).Font
                .color = myColor
            End With
        End If
        
    Next
     
End Sub

Sub S_ColorStringTask2_Repeat()

End Sub

Sub S_ImportAllModules()
'From OpenGPT !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ' Import all modules from workbook1.xlsm
    workbook_path = "C:\path\to\workbook1.xlsm"
    Import "C:\path\to\workbook1.xlsm", "*"
    
    ' Import all modules from module3.bas (a text file)
    'Import "C:\path\to\module3.bas", "*"
End Sub


Sub S_ColorString2()
    Dim currSelection As range
    Set currSelect = selection
    
    txt1 = currSelect.Areas(1)(1).value
    txt2 = currSelect.Areas(2)(1).value
    n1 = Len(txt1)
    n2 = Len(txt2)
    Dim word_list, sentence_list As range
    
    
    If n1 > n2 Then
        Set sentence_list = currSelect.Areas(1)
        Set word_list = currSelect.Areas(2)
    Else
        Set word_list = currSelect.Areas(1)
        Set sentence_list = currSelect.Areas(2)
    End If
    
    
    For Each curr_word In word_list
        n_word = Len(curr_word.value)
        myColor = curr_word.Font.color
        Call S_ColorStringTask(curr_word, n_word, myColor, sentence_list)
        
    Next
    
    
    
    
    'sentence_list.Interior.color = vbBlue
    'word_list.Interior.color = vbRed
    
    
    
    
    
    
    
    
End Sub
Sub ColorFontWith()
    

End Sub

Sub BoldFontAt()

End Sub

Sub BoldFontWith()

End Sub
