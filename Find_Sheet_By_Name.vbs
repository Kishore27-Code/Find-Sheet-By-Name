Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set xlVbscript = objExcel.WorkBooks.Open("C:\Users\Excel.xlsx")

Find_Sheet = "Sheet1"

For TotalSheetsCount = 1 To xlVbscript.Sheets.Count
vSheet = xlVbscript.Sheets(TotalSheetsCount).Name
If Find_Sheet = vSheet Then
Sht = xlVbscript.Sheets(TotalSheetsCount).Name

MsgBox(TotalSheetsCount)'''''''''''''''Finded Sheets Number'''''''''''''

End If
Next

xlVbscript.save
xlVbscript.Close

objExcel.Quit
set objExcel=nothing