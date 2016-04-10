Const xlDataLabelsShowPercent = 3
Const xlPie = 5

Set oExcel = CreateObject("Excel.Application")

Set ws = CreateObject("WScript.Shell")
pwd = ws.CurrentDirectory

Set oWorkbook = oExcel.Workbooks.Open(pwd + "\test.xlsx")
Set oSheet = oWorkbook.Worksheets(1)

Set oChartObject = oSheet.ChartObjects.Add(250, 30, 600, 400)
oChartObject.Chart.SetSourceData oSheet.Range("A1", "B13")
oChartObject.Chart.ChartType = xlPie
oChartObject.Chart.HasTitle = false
oChartObject.Chart.ApplyDataLabels xlDataLabelsShowPercent

oWorkbook.Save
oWorkbook.Close

MsgBox "done"