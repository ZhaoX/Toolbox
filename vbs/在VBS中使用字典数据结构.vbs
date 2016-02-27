Set oExcel = CreateObject("Excel.Application")

Set ws = CreateObject("WScript.Shell")
pwd = ws.CurrentDirectory

'构建变更原因和变更类型的对照字典
Set oDictionary = CreateObject("Scripting.Dictionary")

Set oMapWorkbook = oExcel.Workbooks.Open(pwd + "\变更原因与变更类型对照关系.xls")
Set oMapSheet = oMapWorkbook.Worksheets(1)

Set oMapRange = oMapSheet.UsedRange

For rowIndex = 2 To oMapRange.Rows.count
    If Not(oDictionary.Exists(oMapSheet.Cells(rowIndex, 1).Value)) Then
        oDictionary.Add oMapSheet.Cells(rowIndex, 1).Value, oMapSheet.Cells(rowIndex, 3).Value
    Else
        'MsgBox "重复的变更原因：" + vbCrLf + oMapSheet.Cells(rowIndex, 1).Value
    End If
Next

oMapWorkbook.Close

'按行读取汇总表，找到对应模板并生成相应的excel文件
Set oListWorkbook = oExcel.Workbooks.Open(pwd + "\test20160217.xls")
Set oListSheet = oListWorkbook.Worksheets(1)

Set oListRange = oListWorksheet.UsedRange

For rowIndex =1 To oListRange.Rows.count-2

Next

oListWorkbook.Close

