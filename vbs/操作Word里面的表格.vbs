Const wdReplaceAll = 2

inputPath = GetInputPath()

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(inputPath)

confNo = ExtractConfNo(workbook)
confDay = GetConfDay()

Set word = CreateObject("Word.Application")
Set doc = word.Documents.Open(CurDirectory + "\数据中心CCB会议纪要（模板）.docx")


doc.Content.Find.Execute "[ConfNo]",,,,,,,,,confNo,wdReplaceAll
doc.Content.Find.Execute "[ConfDay]",,,,,,,,,confDay,wdReplaceAll

Set sheetBGSS = workbook.Sheets("变更实施申请汇总")
For rowIndex = 1 To sheetBGSS.UsedRange.Rows.count
    If (sheetBGSS.Cells(rowIndex, "G").Value = "需上会陈述") Then
        doc.Tables(1).Range.Rows.Add
    End If
Next

Set sheetLXBG = workbook.Sheets("例行变更申请汇总")
For rowIndex = 1 To sheetLXBG.UsedRange.Rows.count
    If (sheetLXBG.Cells(rowIndex, "I").Value = "需上会陈述") Then
        doc.Tables(1).Range.Rows.Add
    End If
Next

doc.Tables(1).Range.Rows(doc.Tables(1).Range.Rows.Count).Delete

doc.SaveAs(CurDirectory + "\数据中心CCB会议纪要" + confNo + "期.docx")
doc.Close
Set doc = word.Documents.Open(CurDirectory + "\数据中心CCB会议纪要" + confNo + "期.docx")

wordTableRowIndex=2

For rowIndex = 1 To sheetBGSS.UsedRange.Rows.count
    If (sheetBGSS.Cells(rowIndex, "G").Value = "需上会陈述") Then
        doc.Tables(1).cell(wordTableRowIndex, 2).Range.Text = sheetBGSS.Cells(rowIndex, "C").Value
        doc.Tables(1).cell(wordTableRowIndex, 3).Range.Text = sheetBGSS.Cells(rowIndex, "F").Value
        doc.Tables(1).cell(wordTableRowIndex, 4).Range.Text = sheetBGSS.Cells(rowIndex, "I").Value + vbCrLf + sheetBGSS.Cells(rowIndex, "H").Value
        wordTableRowIndex = wordTableRowIndex + 1
    End If
Next

For rowIndex = 1 To sheetLXBG.UsedRange.Rows.count
    If (sheetLXBG.Cells(rowIndex, "I").Value = "需上会陈述") Then
        doc.Tables(1).cell(wordTableRowIndex, 2).Range.Text = sheetLXBG.Cells(rowIndex, "E").Value
        doc.Tables(1).cell(wordTableRowIndex, 3).Range.Text = sheetLXBG.Cells(rowIndex, "D").Value
        doc.Tables(1).cell(wordTableRowIndex, 4).Range.Text = sheetLXBG.Cells(rowIndex, "C").Value
        wordTableRowIndex = wordTableRowIndex + 1
    End If
Next

doc.Save
doc.Close
word.Quit

workbook.Close
excel.Quit
MsgBox "Done!"

'-----------------------------------------------------------------------------------------------------
' Internal Functions
'-----------------------------------------------------------------------------------------------------
Function GetInputPath()
    Set wShell=CreateObject("WScript.Shell")
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    GetInputPath = oExec.StdOut.ReadLine
End Function

Function ExtractConfNo(workbook)
    Set sheet = workbook.Sheets("变更实施申请汇总")
    val1 = sheet.Cells(1, "A").Value
    val2 = Replace(val1, "2018", "")
    ExtractConfNo = CleanString(val2)
End Function

Function CleanString(strIn)
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[^\d]+"
    CleanString = .Replace(strIn, vbNullString)
    End With
End Function

Function CurDirectory()
    Set WshShell = WScript.CreateObject("WScript.Shell")
    CurDirectory = WshShell.CurrentDirectory
End Function

Function GetConfDay()
    GetConfDay = CStr(Year(Now)) + "年" + CStr(Month(Now)) + "月" + CStr(Day(Now)) + "日"
End Function