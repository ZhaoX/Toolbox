Const FOR_READING = 1
strFilePath = "C:\Users\zhaox\Desktop\tiku1.txt"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile(strFilePath, FOR_READING)

strExcelPath = "C:\Users\zhaox\Desktop\tiku.xlsx"
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add
intColumn = 1

strQuestion = ""
strCheckForString = "Answer"

Do Until objFile.AtEndOfStream
    strCurLine = objFile.ReadLine
    If (Left(LTrim(strCurLine),Len(strCheckForString)) = strCheckForString) Then  
       'MsgBox strQuestion
       objExcel.Cells(intColumn, 1).Value = strQuestion
       objExcel.Cells(intColumn, 2).Value = Mid(strCurLine, 8, 10)
       strQuestion = ""
       intColumn = intColumn + 1
    ElseIf (strCurLine = "The safer , easier way to help you pass any IT exams.") Then
       strQuestion = strQuestion
    ElseIf (Right(strCurLine, 5) = "/ 193") Then
       strQuestion = strQuestion
    ElseIf Not((Asc(LTrim(strCurLine)) > 64 and Asc(LTrim(strCurLine)) < 91)) Then
       strQuestion = strQuestion + " " + strCurLine
    Else
       strQuestion = strQuestion + vbCrLf + strCurLine
    End If
Loop

objExcel.Workbooks(1).SaveAs(strExcelPath)

objExcel.Workbooks.Close
objFile.Close