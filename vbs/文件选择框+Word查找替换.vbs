MsgBox "请选择会议审批汇总表"
inputPath = GetInputPath()

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(inputPath)

Set sheet2 = workbook.Sheets("变更实施申请汇总")
Set sheet3 = workbook.Sheets("例行变更申请汇总")
Set sheet4 = workbook.Sheets("生产变更排程")

S2 = 0
S2Z = 0
S2P = 0
S2H = 0
S2HZ = 0
S2HP = 0
S3 = sheet3.UsedRange.Rows.count - 2
S4 = sheet4.UsedRange.Rows.count - 2

For rowIndex = 1 To sheet2.UsedRange.Rows.count
    If (sheet2.Cells(rowIndex, "D").Value = "重要变更") Then
        S2Z = S2Z + 1
        If (sheet2.Cells(rowIndex, "G").Value = "需上会陈述") Then
            S2HZ = S2HZ + 1        
        End If
    End If
    
    If (sheet2.Cells(rowIndex, "D").Value = "普通变更") Then
        S2P = S2P + 1
        If (sheet2.Cells(rowIndex, "G").Value = "需上会陈述") Then
            S2HP = S2HP + 1        
        End If
    End If
Next

S2 = S2Z + S2P
S2H = S2HZ + S2HP

workbook.Close
excel.Quit

MsgBox "请选择需要更新的会议纪要"
inputPath = GetInputPath()
Set word = CreateObject("Word.Application")
Set doc = word.Documents.Open(inputPath)

Const wdReplaceAll = 2
doc.Content.Find.Execute "[S4+S3]",,,,,,,,,CStr(S4 + S3),wdReplaceAll
doc.Content.Find.Execute "[S3]",,,,,,,,,CStr(S3),wdReplaceAll
doc.Content.Find.Execute "[S2]",,,,,,,,,CStr(S2),wdReplaceAll
doc.Content.Find.Execute "[S2-Z]",,,,,,,,,CStr(S2Z),wdReplaceAll
doc.Content.Find.Execute "[S2-P]",,,,,,,,,CStr(S2P),wdReplaceAll
doc.Content.Find.Execute "[S2-H]",,,,,,,,,CStr(S2H),wdReplaceAll
doc.Content.Find.Execute "[S2-H-Z]",,,,,,,,,CStr(S2HZ),wdReplaceAll
doc.Content.Find.Execute "[S2-H-P]",,,,,,,,,CStr(S2HP),wdReplaceAll
doc.Content.Find.Execute "[S4+S3+S2-(S2-H)]",,,,,,,,,CStr(S4 + S3 + S2 - S2H),wdReplaceAll

doc.Save
doc.Close
word.Quit

MsgBox "Done!"

'-----------------------------------------------------------------------------------------------------
' Internal Functions
'-----------------------------------------------------------------------------------------------------
Function GetInputPath()
    Set wShell=CreateObject("WScript.Shell")
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    GetInputPath = oExec.StdOut.ReadLine
End Function