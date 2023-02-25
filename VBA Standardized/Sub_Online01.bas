Attribute VB_Name = "Sub_Online01"

Sub CreateAndFillExcelFile()
'From Chat GPT
'Fill in values without opening Excel File
  Dim xlApp As Excel.Application
  Dim xlWorkbook As Excel.Workbook
  Dim xlWorksheet As Excel.Worksheet

  Set xlApp = New Excel.Application
  Set xlWorkbook = xlApp.Workbooks.Add
  Set xlWorksheet = xlWorkbook.Sheets.Add

  xlWorksheet.Cells(1, 1).value = "Header 1"
  xlWorksheet.Cells(1, 2).value = "Header 2"
  xlWorksheet.Cells(2, 1).value = "Value 1"
  xlWorksheet.Cells(2, 2).value = "Value 2"

  xlWorkbook.SaveAs "C:\Users\Heng2020\OneDrive\W_Documents\Rotation 2 EastProduct.xlsx"
  xlWorkbook.Close
  xlApp.Quit
End Sub
'https://www.youtube.com/watch?v=h_sC6Uwtwxk&ab_channel=LeilaGharani
'Select file in Excel Leila Gharani
Sub SelectFile()
    Dim filepath As Variant
    filepath = Application.GetOpenFilename(Title:="Browse your file", FileFilter:="Excel Files (*.xls*),*xls*")
    Application.ScreenUpdating = False
    'To prevent Excel file from flickering when open new file
    Set xlApp = New Excel.Application
    If filepath <> False Then
        Set xlWorkbook = xlApp.Workbooks.Open(filepath)
    End If
    Application.ScreenUpdating = True
    
End Sub


