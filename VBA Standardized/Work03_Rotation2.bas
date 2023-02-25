Attribute VB_Name = "Work03_Rotation2"
Function Rg_NextNoTextCell(rng, direction As XlDirection) As range
    ' Declare a variable to store the next blank cell
    
    Dim nextBlankCell As range
    Dim row_offset As Integer
    Dim col_offset As Integer

    ' Set the next blank cell to the current range
    Set nextBlankCell = rng
    ' Set the offset
    If direction = xlUp Then
        row_offset = -1
        col_offset = 0
    ElseIf direction = xlDown Then
        row_offset = 1
        col_offset = 0
    ElseIf direction = xlToLeft Then
        row_offset = 0
        col_offset = -1
    ElseIf direction = xlToRight Then
        row_offset = 0
        col_offset = 1
    End If
    ' Loop until a blank cell is found or the end of the worksheet is reached
    Do Until IsEmpty(nextBlankCell) Or nextBlankCell Is Nothing
        ' Move to the next cell in the specified direction
        Set nextBlankCell = nextBlankCell.Offset(row_offset, col_offset)
    Loop
    
    ' Return the next blank cell
    Set Rg_NextNoTextCell = nextBlankCell


End Function

Function Rg_ReorderRange(rng)
'Reorder range by their addresses when it's not in order
    ws_name = rng.Parent.name
    Dim ws As Worksheet
    Set ws = Worksheets(ws_name)
    
    Dim out_rng As range
    Dim addresses() As String
    ReDim addresses(0 To rng.Cells.count - 1) As String
    Dim i As Long
    i = 0
    For Each cell In rng
        addresses(i) = cell.Address
        i = i + 1
    Next cell
    ' sort the addresses array
    Dim temp As String
    For i = 0 To UBound(addresses) - 1
        For j = i + 1 To UBound(addresses)
            If addresses(i) > addresses(j) Then
                temp = addresses(i)
                addresses(i) = addresses(j)
                addresses(j) = temp
            End If
        Next j
    Next i
    Set out_rng = ws.range(addresses(0))
    For i = 1 To UBound(addresses)
        Set out_rng = Union(out_rng, ws.range(addresses(i)))
    Next i
    Set Rg_ReorderRange = out_rng
End Function

Function Rg_FindSomeRanges(searchString As Variant, inx_arr, Optional ws = "")
'searchString: array or string
'Return the ranges with only some index
'eg I want only 2nd then inx_arr = 1
'If I want 2nd,3rd and 5th then inx = [1,2,4]
End Function
Function Rg_FindAllRange(searchString As Variant, Optional ws = "") As range
'Hard for ChatGPT
    Dim targetSheet As Worksheet
        If ws = "" Then
            ws = ActiveSheet.name
        End If
    
    If TypeName(ws) = "String" Then
        Set targetSheet = ThisWorkbook.Sheets(ws)
    ElseIf TypeName(ws) = "Worksheet" Then
        Set targetSheet = ws
    Else
        MsgBox "Invalid input for worksheet"
        Exit Function
    End If
    
    Dim outRange As range
    
    If IsArray(searchString) Then
        For i = LBound(searchString) To UBound(searchString)
            Set currFound = Rg_FindAllRangeH1(searchString(i), ws)
            Set outRange = Rg_Union(outRange, currFound)
        Next
    Else
        Set outRange = Rg_FindAllRangeH1(searchString, ws)
    End If
    
    Set Rg_FindAllRange = outRange
End Function

Function Rg_FindAllRangeH1(searchString, Optional ws = "")
'It works
'Hard for ChatGPT
    Dim targetSheet As Worksheet
        If ws = "" Then
            ws = ActiveSheet.name
        End If
    
    If TypeName(ws) = "String" Then
        Set targetSheet = ThisWorkbook.Sheets(ws)
    ElseIf TypeName(ws) = "Worksheet" Then
        Set targetSheet = ws
    Else
        MsgBox "Invalid input for worksheet"
        Exit Function
    End If
    Dim SearchRange As range
    Set SearchRange = targetSheet.usedRange ' adjust the range to your needs

    Dim foundRange As range
    Set foundRange = SearchRange.Find(searchString, LookIn:=xlValues)

    Set resFind0 = SearchRange.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart)
    Set curr_found = resFind0
    Dim allFoundRanges As range
    
    If resFind0 Is Nothing Then
        Set Rg_FindAllRangeH1 = Nothing
        Exit Function
    Else
        Set allFoundRanges = resFind0
    End If
    first_address = resFind0.Address
    
    
    Do While True
        Set curr_found = SearchRange.Find(What:=searchString, after:=curr_found, LookIn:=xlValues, LookAt:=xlPart)
        found_sheet_name = curr_found.Parent.name
        found_address = curr_found.Address
        saved_address = found_sheet_name & "   " & found_address
        
        
        If curr_found.Address = first_address Then
            Set Rg_FindAllRangeH1 = allFoundRanges
            Exit Do
        Else
            Set allFoundRanges = Union(allFoundRanges, curr_found)
        End If
       
    Loop
    
    
End Function
Function St_FileNameFromPath(filepath As String) As String
    Dim fileName As String
    fileName = Split(filepath, "\")(UBound(Split(filepath, "\")))
    St_FileNameFromPath = fileName
End Function

Function W_SMCode(prodCode)
    Dim ws01 As Worksheet
    Dim prodCodeList, SMCode As range
    Set ws01 = ThisWorkbook.Sheets("Master")
    Set prodCodeList = ws01.range("B:B")
    Set SMCode = ws01.range("D:D")
    If IsNumeric(prodCode) Then
        prodCode = CLng(prodCode)
    End If
    res = VB_Xlookup(prodCode, prodCodeList, SMCode)
    W_SMCode = res
    
    
End Function
Sub Wb_ColorSheetTabsAfter(wb, inx, Optional tab_color = -0.5)
'Create the similar Sub with Color After some Name
'Before some index
'Mid index(between)
'Right
'Or between 2 sheets names(in string)
' Or delete tab color
    If tab_color = -0.5 Then
        tab_color = RGB(255, 153, 0)
    End If
    For i = 1 To wb.Sheets.count
        If i >= inx Then
            wb.Sheets(i).Tab.color = tab_color 'Orange color
        End If
    Next i
End Sub

Sub Wb_CheckNumSheets(wb, threshold As Integer)
    If wb.Sheets.count >= threshold Then
        MsgBox Workbook.name & " has " & wb.Sheets.count & " sheets."
    End If
End Sub

Function OS_OpenExcelLatest(folderPath)
 '--- change this to sub called OS_GetLatestExcelFile
    mostRecentFile = ""
    mostRecentFileDate = #1/1/1900#
    Dim fso, folder, File As Object
    Dim wb_result As Workbook
    ' Loop through the files in the folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each File In folder.Files
        If LCase(Right(File.name, 4)) = ".xls" Or LCase(Right(File.name, 5)) = ".xlsx" Or LCase(Right(File.name, 5)) = ".xlsb" Or LCase(Right(File.name, 5)) = ".xlsm" Then
            If File.DateLastModified > mostRecentFileDate Then
                mostRecentFile = File.path
                mostRecentFileDate = File.DateLastModified
            End If
        End If
    Next File
    
    ' Open the most recent file
    On Error GoTo Err01
    'Don't update links when open
    Set wb_result = Workbooks.Open(fileName:=mostRecentFile, UpdateLinks:=0)
    On Error GoTo 0
    Set OS_OpenExcelLatest = wb_result
    Exit Function
Err01:
    'myFileName = St_FileNameFromPath(folderPath)
    'MsgBox (myFileName & "Has no Excel File")
    MsgBox ("Has no Excel File")
    'OS_OpenExcelLatest = "NoExcelFile"
    Resume Next
    Exit Function
    
End Function
Function St_GetBefore(str, chr_list As Variant)
    outStr = str
    If IsArray(chr_list) Then
        For i = LBound(chr_list) To UBound(chr_list)
            outStr = St_GetBeforeH1(outStr, chr_list(i))
        Next
        St_GetBefore = outStr
    Else
        St_GetBefore = St_GetBeforeH1(str, chr_list)
    End If
End Function

Function St_GetBeforeH1(str, chr) As String
    Dim charPos As Long
    charPos = InStr(str, chr)
    If charPos > 0 Then
        St_GetBeforeH1 = Left(str, charPos - 1)
    Else
        St_GetBeforeH1 = str
    End If
End Function


Sub W_GetTemplateExample()
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' 1st Finish at 2 hr 15 min
'Not Done Yet 3 hr + Of Debugging
    Dim fso As Object
    Dim folder As Object
    Dim File As Object
    Dim fileName As String
'folder path of wb_numbers (file that P.Oat fill it manually)
    folderPath = "C:\Users\n1603499\OneDrive - Liberty Mutual\Documents\12.02  Rotation2  EastProduct\Test01"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Dim out_arr() As Variant
    Dim wb_result, wb_campaign, wb_numbers As Workbook
    Dim n_sheetLimit As Integer
'Set high numbers if you wanna run all files
'This limit the # of files opened to debug
    n_file_limit = 1000
    n_sheetLimit = 6
    
    output_folder_path = "C:\Users\n1603499\OneDrive - Liberty Mutual\Documents\12.02  Rotation2  EastProduct\Test01\OutPutFile"
    'Watching index when debugging
    count = 0
    For Each File In folder.Files
    'Watching index when debugging
        count = count + 1
        If count > n_file_limit Then
            Exit For
        End If
        If LCase(Right(File.name, 4)) = ".xls" Or LCase(Right(File.name, 5)) = ".xlsx" Or LCase(Right(File.name, 5)) = ".xlsb" Or LCase(Right(File.name, 5)) = ".xlsm" Then
            delimiter = Array("_", " ")
            prodCode = St_GetBefore(Left(File.name, (Len(File.name) - 5)), delimiter)
            
            SMCode = W_SMCode(prodCode)
            outputName = prodCode & "_" & SMCode
            output_path = output_folder_path & "\" & outputName
            Set wb_result = Workbooks.Add
            'Set wb_result = Workbooks.Open("C:\Users\n1603499\OneDrive - Liberty Mutual\Desktop\VBA LibFile V06.02.xlsb")
 


'#####################Template Part############################

            Set wb_numbers = Workbooks.Open(File.path)
            
            Set sheet01 = wb_numbers.Sheets("Coverage Input")
            Set sheet02 = wb_numbers.Sheets("Net Premium Input")
            Set sheet03 = wb_numbers.Sheets("Eligibility Input")
            
            Application.DisplayAlerts = False
            ' Copy the sheets to the wb_result workbook
            sheet01.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            sheet02.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            sheet03.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            Application.DisplayAlerts = True
            
            ' Change the tab color to green
            For Each ws_temp In wb_result.Sheets
                ws_temp.Tab.color = RGB(0, 255, 0)
            Next
            'wb_result.Sheets(wb_result.Sheets.count).Tab.color = RGB(0, 255, 0)
            'wb_result.Sheets(wb_result.Sheets.count - 1).Tab.color = RGB(0, 255, 0)
            'wb_result.Sheets(wb_result.Sheets.count - 2).Tab.color = RGB(0, 255, 0)
            
'----------------------------------------Template Part-----------------------------------------
'#####################campaign Part############################
            campaign_folder_path = W_FindClaimPath2(prodCode)
            
            Set page01 = wb_result.Worksheets("Sheet1")
            page01.name = "FolderPath"
            page01.range("A1") = campaign_folder_path
            
            On Error GoTo Err01
            Set wb_campaign = OS_OpenExcelLatest(campaign_folder_path)
            
            If wb_campaign.Sheets.count > n_sheetLimit Then
' Check this line
'Add the name to indicate that this product code has too many sheets in campaign files
                output_path = output_path & "_ManySheets"
                Application.DisplayAlerts = False
                wb_result.SaveAs (output_path)
                Application.DisplayAlerts = True
                Call Wb_CheckNumSheets(wb_campaign, n_sheetLimit)
                GoTo SkipIteration
            End If
           'Save as without asking
            Application.DisplayAlerts = False
            wb_result.SaveAs (output_path)
            Application.DisplayAlerts = True
            
            
            oldSheetNum = wb_result.Sheets.count
            ColorOrange = CLng("&H" & "FF6600")
            On Error GoTo 0
            
            wb_result.Activate
            Application.DisplayAlerts = False
            For Each ws In wb_campaign.Sheets
                ws.Copy after:=wb_result.Sheets(wb_result.Sheets.count)
            Next ws
            Call Wb_ColorSheetTabsAfter(wb_result, oldSheetNum + 1, ColorOrange)

            'ThisWorkbook.Activate
SkipIteration:
            wb_numbers.Close SaveChanges:=False
            wb_campaign.Close SaveChanges:=False
            Application.DisplayAlerts = True
            
            wb_result.Close
Back01:
'----------------------------------------campaign Part-----------------------------------------
            ' ############################################################## Get Lastest Campaign Excel File ##################################################################
           
             ' ----------------------------------------------------------------------------- Get Most Recent Campaign Excel File ---------------------------------------------------------------------------------------------
            
            
        End If
    Next File
    Exit Sub
Err01:
    'newFileName = wb_result.name & "_NoExcelCampaignFile"
    output_path = output_path & "_NoExcelCampaignFile"
    Application.DisplayAlerts = False
    wb_result.SaveAs (output_path)
    Application.DisplayAlerts = True
            
    wb_numbers.Close
    wb_result.Close
    Resume Back01
    
End Sub
