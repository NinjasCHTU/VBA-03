Attribute VB_Name = "Func_Online1"
Function O_FindAll(SearchRange, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As range
'This doesn't work
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAll
' This searches the range specified by SearchRange and returns a Range object
' that contains all the cells in which FindWhat was found. The search parameters to
' this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing. If
' BeginsWith is not an empty string, only those cells that begin with BeginWith
' are included in the result. If EndsWith is not an empty string, only those cells
' that end with EndsWith are included in the result. Note that if a cell contains
' a single word that matches either BeginsWith or EndsWith, it is included in the
' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim FoundCell As range
    Dim FirstFound As range
    Dim LastCell As range
    Dim ResultRange As range
    Dim XLookAt As XlLookAt
    Dim include As Boolean
    Dim CompMode As VbCompareMethod
    Dim Area As range
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim BeginB As Boolean
    Dim EndB As Boolean
    
    
    CompMode = BeginEndCompare
    If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
        XLookAt = xlPart
    Else
        XLookAt = LookAt
    End If
    
    ' this loop in Areas is to find the last cell
    ' of all the areas. That is, the cell whose row
    ' and column are greater than or equal to any cell
    ' in any Area.
    
    For Each Area In SearchRange.Areas
        With Area
            If .Cells(.Cells.count).row > MaxRow Then
                MaxRow = .Cells(.Cells.count).row
            End If
            If .Cells(.Cells.count).column > MaxCol Then
                MaxCol = .Cells(.Cells.count).column
            End If
        End With
    Next Area
    Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)
    
    On Error GoTo 0
    Set FoundCell = SearchRange.Find(What:=FindWhat, _
            after:=LastCell, _
            LookIn:=LookIn, _
            LookAt:=XLookAt, _
            SearchOrder:=SearchOrder, _
            MatchCase:=MatchCase)
    
    If Not FoundCell Is Nothing Then
        Set FirstFound = FoundCell
        Do Until False ' Loop forever. We'll "Exit Do" when necessary.
            include = False
            If BeginsWith = vbNullString And EndsWith = vbNullString Then
                include = True
            Else
                If BeginsWith <> vbNullString Then
                    If StrComp(Left(FoundCell.text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                        include = True
                    End If
                End If
                If EndsWith <> vbNullString Then
                    If StrComp(Right(FoundCell.text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                        include = True
                    End If
                End If
            End If
            If include = True Then
                If ResultRange Is Nothing Then
                    Set ResultRange = FoundCell
                Else
                    Set ResultRange = Application.Union(ResultRange, FoundCell)
                End If
            End If
            Set FoundCell = SearchRange.FindNext(after:=FoundCell)
            If (FoundCell Is Nothing) Then
                Exit Do
            End If
            If (FoundCell.Address = FirstFound.Address) Then
                Exit Do
            End If
    
        Loop
    End If
        
    Set O_FindAll = ResultRange

End Function

Function O_ThisWSName()
    'https://stackoverflow.com/questions/19323343/get-a-worksheet-name-using-excel-vba
    O_ThisWSName = Application.Caller.Worksheet.name
    
End Function

Function O_GetSheetName(Optional output_option = 0)
'output_option = 0 Vertical
'output_option = 1 Horizontal (as Array)
'Array1
'WS_Func => Transpose
    Dim outArr() As Variant
    Dim curr_wb As Workbook
    Set curr_wb = Application.Caller.Parent.Parent
    For Each curr_ws In curr_wb.Worksheets
        outArr = A_Append(outArr, curr_ws.name)
    Next
    If output_option = 1 Then
    
        O_GetSheetName = outArr
    Else
        O_GetSheetName = VB_Transpose(outArr)
    End If
    
End Function


' https://stackoverflow.com/questions/4734794/how-to-search-for-a-string-in-all-sheets-of-an-excel-workbook
' Dim sheetCount As Integer
' Dim datatoFind

' Sub Button1_Click()

'     O_Find_Data

' End Sub

' Sub O_Find_Data()
'     Dim counter As Integer
'     Dim currentSheet As Integer
'     Dim notFound As Boolean
'     Dim yesNo As String

'     notFound = True

'     On Error Resume Next
'     currentSheet = ActiveSheet.Index
'     datatoFind = StrConv(InputBox("Please enter the value to search for"), vbLowerCase)
'     If datatoFind = "" Then Exit Sub
'     sheetCount = ActiveWorkbook.Sheets.Count
'     If IsError(CDbl(datatoFind)) = False Then datatoFind = CDbl(datatoFind)
'     For counter = 1 To sheetCount
'         Sheets(counter).Activate

'         Cells.Find(what:=datatoFind, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
'         :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'         False, SearchFormat:=False).Activate

'         If InStr(1, StrConv(ActiveCell.Value, vbLowerCase), datatoFind) Then
'             notFound = False
'             If O_HasMoreValues(counter) Then
'                 yesNo = MsgBox("Do you want to continue search?", vbYesNo)
'                 If yesNo = vbNo Then
'                     Sheets(counter).Activate
'                     Exit For
'                 End If
'             Else
'                 Sheets(counter).Activate
'                 Exit For
'             End If
'             Sheets(counter).Activate
'         End If
'     Next counter
'     If notFound Then
'         MsgBox ("Value not found")
'         Sheets(currentSheet).Activate
'     End If
' End Sub

' Private Function O_HasMoreValues(ByVal sheetCounter As Integer) As Boolean
'     O_HasMoreValues = False
'     Dim str As String
'     Dim lastRow As Long
'     Dim lastCol As Long
'     Dim rRng  As Excel.Range

'     For counter = sheetCounter + 1 To sheetCount
'         Sheets(counter).Activate

'         lastRow = ActiveCell.SpecialCells(xlLastCell).row
'         lastCol = ActiveCell.SpecialCells(xlLastCell).column

'         For vRow = 1 To lastRow
'             For vCol = 1 To lastCol
'                 str = Sheets(counter).Cells(vRow, vCol).text
'                 If InStr(1, StrConv(str, vbLowerCase), datatoFind) Then
'                     O_HasMoreValues = True
'                     Exit For
'                 End If
'             Next vCol

'             If O_HasMoreValues Then
'                 Exit For
'             End If
'         Next vRow

'         If O_HasMoreValues Then
'             Sheets(sheetCounter).Activate
'             Exit For
'         End If
'     Next counter
' End Function




 
 
Function O_isItInThisWB(sLookFor As String)
    For Each curr_ws In ThisWorkbook
        ws_name02 = curr_ws.name
        ' str_find = O_isItInWS(ws_name02, sLookFor)
        ' If str_find <> "Not Found" Then
        '     res_txt = ws_name02 & "***" & str_find
        '     O_isItInThisWB = res_txt
        '     End Function
        ' End If
    Next
    O_isItInThisWB = "Not Found"
    
End Function
 


    

'https://myengineeringworld.net/2013/07/add-description-to-custom-vba-function.html
Sub AddFunctionDescription()

    '------------------------------------------------------------------------
    'This sub can add a description to a selected user-defined function,
    '(UDF) as well as to its parameters, by using the MacroOptions method.
    'After running successfully the macro the UDF function no longer appears
    'to the UDF category of functions, but into the desired category.
    
    'By Christos Samaras
    'Date: 23/07/2013
    'xristos.samaras@gmail.com
    'https://myengineeringworld.net/////
    '------------------------------------------------------------------------
    
    'Delclaring the necessary variables
    Dim FuncName As String
    Dim FuncDesc As String
    Dim FuncCat As Variant
    
    'Depending on the function arguments define the necessary variables on the arry.
    'Here UDF funciton has four arguments, so four variables are declared.
    Dim ArgDesc(1 To 3) As String
    
    '"FrictionFactor" is the name of the function.
    FuncName = "Sp_SelectSkipVB"

    
    'Here we add the function's description.
    FuncDesc = "The generalized select skip.    " & _
    "If you want to select 1 line skip 1 line then r should be 0 or 1  and n =2"

    
    'Choose the built-in function category (it will no longer appear in UDF category).
    'For example, 15 is the engineering category, 4 is the statistical category etc.
    'See the code at the end for all available categories.
    'FuncCat = 15
    
    'You can also use instead of numbers the full category name, for example:
    'FuncCat = "Engineering"
    'Or you can define your own custom category:
    'FuncCat = "My VBA Functions"
    
    'Here we add the description for the function's arguments.
    ArgDesc(1) = "array of items"
    ArgDesc(2) = " is the remander selected r must be less than n"
    ArgDesc(3) = " is the # of repeated cycles"
    

    'Using the MacroOptions method add the function description (and its arguments).
    Application.MacroOptions _
        Macro:=FuncName, _
        Description:=FuncDesc, _
        ArgumentDescriptions:=ArgDesc
    
    'Category:=FuncCat, _
    'Available built-in categories in Excel.
    'This select case is somehow irrelevelant, but it was added for
    'demonstration purposues.
    Select Case FuncCat
        Case 1: FuncCat = "Financial"
        Case 2: FuncCat = "Date & Time"
        Case 3: FuncCat = "Math & Trig"
        Case 4: FuncCat = "Statistical"
        Case 5: FuncCat = "Lookup & Reference"
        Case 6: FuncCat = "Database"
        Case 7: FuncCat = "Text"
        Case 8: FuncCat = "Logical"
        Case 9: FuncCat = "Information"
        Case 10: FuncCat = "Commands"
        Case 11: FuncCat = "Customizing"
        Case 12: FuncCat = "Macro Control"
        Case 13: FuncCat = "DDE/External"
        Case 14: FuncCat = "User Defined default"
        Case 15: FuncCat = "Engineering"
        Case Else: FuncCat = FuncCat
    End Select

    'Inform the user about the process.
    MsgBox FuncName & " was successfully added to the " & FuncCat & " category!", vbInformation, "Done"
    
End Sub



'redim preserve both dimensions for a multidimension array *ONLY
'https://newbedev.com/excel-vba-how-to-redim-a-2d-array
Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
    ReDimPreserve = False
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function

Sub CopyLambdas()
'https://stackoverflow.com/questions/69872165/how-to-share-generic-lambda-functions-over-different-projects
'Copy Lambda Function to any Excel that open
'NEED improvement
'Right now to use it you must have a sheet named 'Lambdas' in your lambdaFile
'Then all other files you must delete your copied sheet manually
    Dim wb As Workbook, n, List
    'make a concatenated list of lambdas in this workbook
    List = "|"                                   'delimiter is |
    For Each n In ThisWorkbook.Names
        If InStr(1, n.value, "lambda", vbTextCompare) > 0 Then
            List = List & n.name & "|"
        End If
    Next n
       
    'process all open workbooks (except this one of course)
    For Each wb In Workbooks
        If Not wb Is ThisWorkbook Then
            With wb
                For Each n In .Names             'look for lambdas
                    If InStr(1, n.value, "lambda", vbTextCompare) > 0 Then
                        'if this lambda has a name that's in our list, delete it
                        If InStr(1, "|" & n.name & "|", n.name, vbTextCompare) > 0 Then n.Delete
                    End If
                Next n
                ThisWorkbook.Sheets("Lambdas").Copy after:=.Sheets(.Sheets.count)
            End With
        End If
    Next wb
    MsgBox ("Lambda Function Transfer Completed")
End Sub



