targetDate = "2017-06-22"

'����ָ�����ڵĹ����б�
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
oXMLHTTP.Open "GET", DownloadUrl(targetDate), False
oXMLHTTP.Send

Set oFs = CreateObject("Scripting.FileSystemObject")
If oFs.FileExists(DownloadFileName(targetDate)) Then
    oFs.DeleteFile DownloadFileName(targetDate)
End If

If oXMLHTTP.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write oXMLHTTP.responseBody
    oStream.SaveToFile DownloadFileName(targetDate)
    oStream.Close
End If

'ͳ����ѯ����Ͷ����������������
Set oExcel = CreateObject("Excel.Application")
Set ws = CreateObject("WScript.Shell")
pwd = ws.CurrentDirectory

consult = 0
complaint = 0
followUp = 0

Set oIssuesWorkbook = oExcel.Workbooks.Open(pwd + "\" + DownloadFileName(targetDate))
Set oIssuesSheet = oIssuesWorkbook.Worksheets(1)
Set oIssuesRange = oIssuesSheet.UsedRange
For rowIndex = 2 To oIssuesRange.Rows.count
    If oIssuesSheet.Cells(rowIndex, "J") = "�˺�����" And oIssuesSheet.Cells(rowIndex, "K") = "VIP����" Then
        If oIssuesSheet.Cells(rowIndex, "H") = "��ѯ" Then
            consult = consult + 1
        Else
            complaint = complaint + 1
        End If
        
        If oIssuesSheet.Cells(rowIndex, "P") = "��" Then
            followUp = followUp + 1
        End If
    End If
Next
oIssuesWorkbook.Close

'����ָ�������ձ�
Set oDailyWorkbook = oExcel.Workbooks.Open(pwd + "\" + "DailyReportTemplate.xlsx")
Set oDailySheet = oDailyWorkbook.Worksheets(1)
oDailySheet.Cells(2, "A") = consult
oDailySheet.Cells(2, "B") = complaint
oDailySheet.Cells(2, "C") = followUp
oDailyWorkbook.SaveAs pwd+"\"+DailyReportFileName(targetDate)
oDailyWorkbook.Close

MsgBox CStr(consult) + vbCrLf + CStr(complaint) + vbCrLf + CStr(followUp) 

'--------------------------------------------------------------------------------------
'Internal Functions
'--------------------------------------------------------------------------------------
Function DownloadUrl(targetDate)
    DownloadUrl = "http://10.221.83.175/data/invoke/" + Replace(targetDate, "-", "/", 1, 1) + "/" + targetDate + ".csv"
End Function

Function DownloadFileName(targetDate)
    DownloadFileName = targetDate + ".csv"
End Function

Function DailyReportFileName(targetDate)
    DailyReportFileName = "PASSPORT�û�Ͷ���ܽ�-" + Replace(targetDate, "-", "") + ".xlsx"
End Function