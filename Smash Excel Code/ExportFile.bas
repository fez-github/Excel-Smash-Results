Attribute VB_Name = "ExportFile"
Sub Export(shtName As String)
'Purpose: To make file compatible with Google sheets.  Adjusts certain formulas to values, then saves as XLSX.
'Last Updated: September 27, 2021

'Disable alerts
    Application.DisplayAlerts = False

'Delete all unnecessary sheets
    Sheets("Userform").Delete
    Dim i As Integer
    Dim FirstSheet As Integer
        FirstSheet = 2
    Dim LastSheet As Integer
        LastSheet = Worksheets.Count
        
    For i = FirstSheet To LastSheet
        Sheets(i).Cells.Copy
        Sheets(i).Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Next i
'Save workbook without macros.
    ThisWorkbook.SaveAs Filename:=shtName, FileFormat:=xlWorkbookDefault
End Sub
