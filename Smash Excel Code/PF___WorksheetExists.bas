Attribute VB_Name = "PF_WorksheetExists"
Public Function WorksheetExists(wb As Workbook, shtName As String) As Boolean
'Purpose: To see if a worksheet name is already used before adding a new one.
'Last Updated: July 26, 2021

'Declare variables
    Dim sht As Worksheet
'Search through each worksheet in the workbook for the given sheet name.
    For Each sht In wb.Sheets
        If sht.name = shtName Then
            WorksheetExists = True
            Exit Function
        End If
    Next sht
'Return False if name was not found.
    WorksheetExists = False
End Function

