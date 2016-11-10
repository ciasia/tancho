if WScript.Arguments.Count < 1 Then
    WScript.Echo "Please provide a file"
    Wscript.Quit
End If

Dim excel_obj
Set excel_obj = CreateObject("Excel.Application")
Dim workbook_obj
Set workbook_obj = excel_obj.Workbooks.Open(Wscript.Arguments.Item(0))

For Each sheet_obj In workbook_obj.sheets
	sheet_obj.Copy
	excel_obj.ActiveWorkbook.SaveAs workbook_obj.Path & "\" & sheet_obj.Name & ".csv", 6
	excel_obj.ActiveWorkbook.Close False
Next

workbook_obj.Close False
excel_obj.Quit
Set excel_obj = Nothing