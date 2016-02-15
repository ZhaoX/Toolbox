Set excelApp = CreateObject("Excel.Application")
Set totalWorkbook = excelApp.Workbooks.Open("D:\total.xlsx")
Set totalSheet = totalWorkbook.Sheets(4)

For rowIndex = 1 To 5
	'read info from total excel 
	name = totalSheet.Cells(rowIndex, 1).Value
	plan = totalSheet.Cells(rowIndex, 2).Value
	mDate = Mid(totalSheet.Cells(rowIndex, 5).Value, 1, 10)
	mYear = Mid(mDate, 1, 4)
	mMonth = Mid(mDate, 6, 2)
	mDay = Mid(mDate, 9, 2)
	system = totalSheet.Cells(rowIndex, 9).Value
	department = totalSheet.Cells(rowIndex, 10).Value
	sysPerson = totalSheet.Cells(rowIndex, 6).Value
	appPerson = totalSheet.Cells(rowIndex, 11).Value
	
	'generate excel
	Set curWorkbook = excelApp.Workbooks.Open("D:\test.xls")
	Set curSheet = curWorkbook.Sheets(1)
	
	originSubHead = curSheet.Cells(2, 1).Value
	subHead1 = Replace(originSubHead, "[year]", mYear)
	subHead2 = Replace(subHead1, "[month]", mMonth)
	subHead3 = Replace(subHead2, "[day]", mDay)
	subHead4 = Replace(subHead3, "[name]", name)
	curSheet.Cells(2, 1).Value = subHead4
	
	For innerRow = 4 To 11
	    curSheet.Cells(innerRow, 4) = mDate
		curSheet.Cells(innerRow, 8) = "系统管理二 " + sysPerson
	Next
	
	curSheet.Cells(7, 9) = department + " " + appPerson
	curSheet.Cells(9, 9) = department + " " + appPerson
	curSheet.Cells(11, 9) = department + " " + appPerson
	
	curWorkbook.SaveAs("D:\数据中心应急演练记录表-2014年版-" + system + ".xls")
	curWorkbook.Close()
	
	MsgBox "Generated D:\数据中心应急演练记录表-2014年版-" + system + ".xls"
Next

totalWorkbook.Close()
excelApp.Quit
MsgBox "Done!"