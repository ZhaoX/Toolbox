targetDates = Array("2017-06-07", "2017-06-08", "2017-06-09", "2017-06-10", "2017-06-11", "2017-06-12", "2017-06-13")

Set oExcel = CreateObject("Excel.Application")
Set ws = CreateObject("WScript.Shell")
pwd = ws.CurrentDirectory

Set oWeeklyWorkbook = oExcel.Workbooks.Open(pwd + "\" + "WeeklyReportTemplate.xlsx")
Set oWeeklyNumberSheet = oWeeklyWorkbook.Worksheets(1)
Set oWeeklyProblemSheet = oWeeklyWorkbook.Worksheets(2)

For index = 0 To 6
    Set oDailyWorkbook = oExcel.Workbooks.Open(DailyReportFilePath(targetDates(index)))
    Set oDailyNumberSheet = oDailyWorkbook.Worksheets(1)
    Set oDailyProblemSheet = oDailyWorkbook.Worksheets(2)
    
    '汇总咨询量、投诉量、后续跟进量
    oWeeklyNumberSheet.Cells(index+2, "A") = Replace(targetDates(index), "-", "/")
    oWeeklyNumberSheet.Cells(index+2, "B") = oDailyNumberSheet.Cells(2, "A").Value
    oWeeklyNumberSheet.Cells(index+2, "C") = oDailyNumberSheet.Cells(2, "B").Value
    oWeeklyNumberSheet.Cells(index+2, "D") = oDailyNumberSheet.Cells(2, "C").Value
    
    '汇总问题列表
    Set oWeeklyProblemRange = oWeeklyProblemSheet.UsedRange
    Set oDailyProblemRange = oDailyProblemSheet.UsedRange
    For rowIndex = 2 To oDailyProblemRange.Rows.count
        oWeeklyProblemSheet.Cells(rowIndex + oWeeklyProblemRange.Rows.count, "A") = oDailyProblemSheet.Cells(rowIndex, "B").Value
        oWeeklyProblemSheet.Cells(rowIndex + oWeeklyProblemRange.Rows.count, "B") = 1
        oWeeklyProblemSheet.Cells(rowIndex + oWeeklyProblemRange.Rows.count, "C") = oDailyProblemSheet.Cells(rowIndex, "C").Value
    Next
	
	oDailyWorkbook.Close
Next

oWeeklyWorkbook.SaveAs pwd+"\"+WeeklyReportFileName(targetDates(0), targetDates(6))
oWeeklyWorkbook.Close

MsgBox "done!"
'--------------------------------------------------------------------------------------
'Internal Functions
'--------------------------------------------------------------------------------------
Function DailyReportFileName(targetDate)
    DailyReportFileName = "PASSPORT用户投诉总结-" + Replace(targetDate, "-", "") + ".xlsx"
End Function

Function DailyReportFilePath(targetDate)
    targetPath = "E:\PASSPORT\用户投诉\" + Left(Replace(targetDate, "-", ""), 6) + "\日报\"
    DailyReportFilePath = targetPath + DailyReportFileName(targetDate)
End Function

Function WeeklyReportFileName(startDate, endDate)
    WeeklyReportFileName = "汇总" + Replace(startDate, "-", "") + "-" + Replace(endDate, "-", "")
End Function