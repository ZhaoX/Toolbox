'此脚本负责对模板进行一些预处理操作

sPath = ".\base"
Set oFso = CreateObject("Scripting.FileSystemObject")  
Set oFolder = oFso.GetFolder(sPath)  
Set oFiles = oFolder.Files  

Set oExcel = CreateObject("Excel.Application")
    
For Each oFile In oFiles  

    Set oWorkbook = oExcel.Workbooks.Open(oFile.Path)
    Set oSheet = oWorkbook.Sheets(1)
    
    '模板里H52的黑山扈都要改成张江
    If (oSheet.Cells(52, "H").value = "黑山扈") Then
        oSheet.Cells(52, "H").value = "张江"
    ElseIf (oSheet.Cells(52, "H").value <> "张江")Then
        MsgBox "该文件的H52既不是黑山扈也不是张江: " + vbCrLf + oFile.Path
    End If
    
    '将后期还需补充的单元格用其他颜色标注 如I8，B13
    osheet.Cells(8, "I").Interior.colorindex = 24
    osheet.Cells(13, "B").Interior.colorindex = 24
    
    oWorkbook.Save
    oWorkbook.Close
Next  

MsgBox "预处理的工作做完啦"