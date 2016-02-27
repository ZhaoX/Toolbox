Set excelApp = CreateObject("Excel.Application")
Set curWorkbook = excelApp.Workbooks.Open("D:\total.xlsx")
Set curSheet = CurWorkbook.Sheets(4)

MsgBox "break point 0"

For rowIndex = 1 To 5
	'read info from sheet 
	name = curSheet.Cells(rowIndex, 1).Value
	plan = curSheet.Cells(rowIndex, 2).Value
	mDate = Mid(curSheet.Cells(rowIndex, 5).Value, 1, 10)
	system = curSheet.Cells(rowIndex, 9).Value
	department = curSheet.Cells(rowIndex, 10).Value
	sysPerson = curSheet.Cells(rowIndex, 6).Value
	appPerson = curSheet.Cells(rowIndex, 11).Value
	
	'generate word file
	Set wordApp = CreateObject("Word.Application")
	Set CurWord = wordApp.Documents.Open("D:\test.doc")
	CurWord.Tables(1).cell(1, 2).Range.Text = name
	CurWord.Tables(1).cell(2, 2).Range.Text = plan 
	CurWord.Tables(1).cell(3, 2).Range.Text = mDate + " 20:00-21:00"
	 
	originSituation = CurWord.Tables(1).cell(6, 2).Range.Text
	situation1 = Replace(originSituation, "[SYSTEM]", system)
	situation2 = Replace(situation1, "[DEPARTMENT]", department)
	CurWord.Tables(1).cell(6, 2).Range.Text = situation2
	
	originResult = CurWord.Tables(1).cell(7, 2).Range.Text
	result1 = Replace(originResult,"[SYS_PERSON]",sysPerson)
	result2 = Replace(result1,"[APP_PERSON]",appPerson)
	CurWord.Tables(1).cell(7, 2).Range.Text = result2
	
	originInfo = CurWord.Tables(1).cell(10, 2).Range.Text
	info1 = Replace(originInfo,"[SYS_PERSON]",sysPerson)
	info2 = Replace(info1, "2014/1/22", mDate)
	CurWord.Tables(1).cell(10, 2).Range.Text = info2
	
	CurWord.SaveAs2("D:\数据中心应急演练评估表-2014年版-" + system + ".doc")
	CurWord.Close()
	wordApp.Quit
	
	MsgBox "Generated D:\数据中心应急演练评估表-2014年版-" + system + ".doc"
Next

curWorkbook.Close()
excelApp.Quit
MsgBox "Done!"

'--------------------------------------------------------------------------------------
'internal functions
'--------------------------------------------------------------------------------------