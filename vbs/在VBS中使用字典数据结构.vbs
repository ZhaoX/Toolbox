Set oExcel = CreateObject("Excel.Application")

Set ws = CreateObject("WScript.Shell")
pwd = ws.CurrentDirectory

'�������ԭ��ͱ�����͵Ķ����ֵ�
Set oDictionary = CreateObject("Scripting.Dictionary")

Set oMapWorkbook = oExcel.Workbooks.Open(pwd + "\���ԭ���������Ͷ��չ�ϵ.xls")
Set oMapSheet = oMapWorkbook.Worksheets(1)

Set oMapRange = oMapSheet.UsedRange

For rowIndex = 2 To oMapRange.Rows.count
    If Not(oDictionary.Exists(oMapSheet.Cells(rowIndex, 1).Value)) Then
        oDictionary.Add oMapSheet.Cells(rowIndex, 1).Value, oMapSheet.Cells(rowIndex, 3).Value
    Else
        'MsgBox "�ظ��ı��ԭ��" + vbCrLf + oMapSheet.Cells(rowIndex, 1).Value
    End If
Next

oMapWorkbook.Close

'���ж�ȡ���ܱ��ҵ���Ӧģ�岢������Ӧ��excel�ļ�
Set oListWorkbook = oExcel.Workbooks.Open(pwd + "\test20160217.xls")
Set oListSheet = oListWorkbook.Worksheets(1)

Set oListRange = oListWorksheet.UsedRange

For rowIndex =1 To oListRange.Rows.count-2

Next

oListWorkbook.Close

