'Accepts an xls/x file as input and creates .csv files for each sheet in the same directory'
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Please provide a file and output location"
    Wscript.Quit
End If

Dim excel_obj
Set excel_obj = CreateObject("Excel.Application")
Dim workbook_obj
Set workbook_obj = excel_obj.Workbooks.Open(Wscript.Arguments.Item(0))
Dim output_dir
output_dir = Wscript.Arguments.Item(1)

For Each sheet_obj In workbook_obj.sheets
	sheet_obj.Copy
	excel_obj.ActiveWorkbook.SaveAs output_dir & "\" & sheet_obj.Name & ".csv", 6
	excel_obj.ActiveWorkbook.Close False
Next

workbook_obj.Close False
excel_obj.Quit
Set excel_obj = Nothing