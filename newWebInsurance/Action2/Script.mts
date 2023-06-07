Set excelObj = CreateObject("Excel.application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

FilePath = "C:\Users\nanangfaisal\Documents\Unified Functional Testing\demoBorlandInsurance\newWebInsurance\GetData.xlsx"
crntDate = FormatDateTime(Now)
crntDate = Replace(crntDate,":",";")
Dim words, ws


Set objWorkbook = excelObj.Workbooks.Open(FilePath)
Set ExcelSheet = excelObj.ActiveWorkbook.Worksheets("Sheet1")
cRow = excelSheet.Usedrange.Rows.count
cColm = excelSheet.Usedrange.Columns.count
excelObj.Application.Visible = True
print cRow

's  = "1001 nol,|1002 satu,|1003 dua,|1004 tiga"
's  = "1001, John Smith1|1001, John Smith1|1001, John Smith1|1001, John Smith1|1001, John Smith1|"
iRowDataLogin = Parameter("iRowDataLogin") - 2
s = Parameter("iFullName")
js = Len(s) - 1
s = Mid(s,1,js)

words  = Split(s,"|")
index = uBound(words) + 1

Redim arrayRow(index, iRowDataLogin) 'somearray in array

For i = 0 To index - 1
	ws = Split(words(i),",")
	ArrayRow(i,g) = ws(g) '+1 to skip 1st row (header)	
Next

'Write Array as a Row
StartRow = 1 + cRow
StartCol = 1

Set Rng = ExcelSheet.Range(ExcelSheet.Cells(StartRow, StartCol), _
	ExcelSheet.Cells(UBound(ArrayRow, 1) - LBound(ArrayRow, 1) + StartRow, UBound(ArrayRow, 2) - LBound(ArrayRow, 2) + StartCol))
Rng.Value = ArrayRow

excelObj.ActiveWorkbook.Save

excelObj.DisplayAlerts = False
objWorkbook.Close False
TerminateProcess()

Set excelObj = Nothing

Sub TerminateProcess
Dim Process 
	For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'EXCEL.EXE'")
   		Process.Terminate
	Next
End Sub
