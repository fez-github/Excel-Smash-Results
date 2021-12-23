Attribute VB_Name = "PF___Unique_Values"
Public Function Unique_Values(dictionary As Object, firstCell As Range)
'Purpose: To acquire all unique values in a range.
'Written by: Mark Hansen
'Last Updated: June 30 2021

'Declare variables
    Dim lookupRange, lookupCell As Range
    Set dictionary = CreateObject("Scripting.Dictionary")
'Declare lookup range.  Looks for the cell entered, and the cell at the bottom of the column.
    Set lookupRange = Range(firstCell, firstCell.End(xlDown))
'Iterate through lookup range and add unique values.
    For Each lookupCell In lookupRange
        If Not dictionary.Exists(lookupCell.Value) Then
            dictionary.Add lookupCell.Value, lookupCell.Value
        End If
    Next lookupCell
'Set dictionary to function.
Set Unique_Values = dictionary
End Function
Private Sub Unique_Values_Initiator()

Dim arrayDictionary As Object
    Set arrayDictionary = Unique_Values(arrayDictionary, Range("B2"))
End Sub
