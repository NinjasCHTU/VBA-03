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

