'�˽ű������ģ�����һЩԤ�������

sPath = ".\base"
Set oFso = CreateObject("Scripting.FileSystemObject")  
Set oFolder = oFso.GetFolder(sPath)  
Set oFiles = oFolder.Files  

Set oExcel = CreateObject("Excel.Application")
    
For Each oFile In oFiles  

    Set oWorkbook = oExcel.Workbooks.Open(oFile.Path)
    Set oSheet = oWorkbook.Sheets(1)
    
    'ģ����H52�ĺ�ɽ�趼Ҫ�ĳ��Ž�
    If (oSheet.Cells(52, "H").value = "��ɽ��") Then
        oSheet.Cells(52, "H").value = "�Ž�"
    ElseIf (oSheet.Cells(52, "H").value <> "�Ž�")Then
        MsgBox "���ļ���H52�Ȳ��Ǻ�ɽ��Ҳ�����Ž�: " + vbCrLf + oFile.Path
    End If
    
    '�����ڻ��貹��ĵ�Ԫ����������ɫ��ע ��I8��B13
    osheet.Cells(8, "I").Interior.colorindex = 24
    osheet.Cells(13, "B").Interior.colorindex = 24
    
    oWorkbook.Save
    oWorkbook.Close
Next  

MsgBox "Ԥ����Ĺ���������"