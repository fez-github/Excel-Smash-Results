VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BracketSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum sheetType
    bracketGames = 0
    bracketSets = 1
    summaryGames = 2
    summarySets = 3
End Enum

Private brktSht As Worksheet
Private plyrCount As Integer
Private playerDict As Object
Private playerCount As Integer
Private maxColumn As String
'Include Sheet name & change depending on Enum
Property Set Sheet(sht As Worksheet)
    Set brktSht = sht
End Property
Property Get name() As String
    name = brktSht.name
End Property
Private Sub Class_Initialize()
    Set playerDict = CreateObject("Scripting.Dictionary")
End Sub
Public Sub InitiateModule(bracketName As String, summaryType As sheetType)
'Purpose: To allow different modules to access this class module, and run the entire procedure at once.
'Last Updated: December 23, 2021

'Check if bracket sheet was already made.  If so, stop the procedure.
    Dim sheetGenerated As Boolean
        sheetGenerated = NewBracketSheet(Replace(bracketName, "/", "-"), summaryType)
        If sheetGenerated = False Then
            Exit Sub
        End If
    GetBracketPlayers bracketName, summaryType
    FormatScoreCells
    AddFormulas summaryType
    FormatLeftColumn
    MergeTopCell
    AddNames
    AutofillFormats
    RangeA1 (bracketName)
    AddNAValues
    RemoveBlankFormulas
    FinalColumns
    FreezePanes
End Sub
Private Function NewBracketSheet(bracketShtName As String, summaryType As sheetType) As Boolean
'Purpose: To create a new type of bracket sheet based on the Match Records table.
'Last Updated: December 23, 2021

'Check for summary type and adjust name.
    Select Case summaryType
        Case sheetType.bracketGames
            bracketShtName = bracketShtName & "(G)"
        Case sheetType.bracketSets
            bracketShtName = bracketShtName & "(S)"
        Case sheetType.summaryGames
            bracketShtName = "AllBrackets(G)"
        Case sheetType.summarySets
            bracketShtName = "AllBrackets(S)"
    End Select
'Check if Bracket Sheet name already exists.  If so, stop sub and inform user.
    Dim bool As Boolean
        bool = WorksheetExists(ThisWorkbook, bracketShtName)
    If bool = True Then
        MsgBox "This sheet name already exists!  Please use a different name, or check to see if the pre-existing sheet can be deleted."
        NewBracketSheet = False
    End If
'Create new sheet with name
        Set brktSht = ThisWorkbook.Sheets.Add(, Sheets(ThisWorkbook.Sheets.Count))
        brktSht.name = bracketShtName
        brktSht.Cells.Font.name = "Arial"
        brktSht.Cells.Font.Size = 10
        brktSht.Rows.RowHeight = 15.75
'Process successful.  Return true
    NewBracketSheet = True
End Function
Private Sub GetBracketPlayers(bracketName As String, summaryType As sheetType)
'Purpose: To get all players for the bracket from the Placings worksheet.
'Last Updated: September 2021

'Declare variables
    Dim i As Integer, j As Integer
        i = 2
'Clear out the player dictionary, in case it's being used multiple times.
    playerDict.RemoveAll
'Search for the bracket name in the Placings sheet, and pull every name that does not have NA in the bracket name.
    With ThisWorkbook.Sheets("Placings")
        If summaryType = sheetType.summaryGames Or summaryType = sheetType.summarySets Then
            j = 1
        Else
            j = .Rows("1:1").Find(bracketName).Column
        End If
        Do Until IsEmpty(.Cells(i, 1)) = True
            If IsEmpty(.Cells(i, j)) = False And .Cells(i, j).Value <> "NA" Then
                playerDict.Add .Cells(i, 1), .Cells(i, 1)
            End If
            i = i + 1
        Loop
    End With
'Declare the player total count in a module-level variable.
    playerCount = playerDict.Count
'Declare the maximum range needed for all the player names in the bracket sheet.
    maxColumn = Split(Cells(1, (playerCount * 2) + 1).Address, "$")(1)
End Sub
Private Sub FormatScoreCells()
'Purpose: To add formatting for the cells used to store scores.
'Last Updated: July 2021
    With brktSht.Range("B2:C2")
        .Borders(xlEdgeBottom).Color = -16777216
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Color = -2500135
        .Borders(xlEdgeRight).Weight = xlThick
        .HorizontalAlignment = xlHAlignCenter
    End With
End Sub
Private Sub AddFormulas(summaryType As sheetType)
'Purpose: Insert formulas into the first 2 cells.  These will be autofilled later.
'Last Updated: July 2021
With brktSht
    If summaryType = bracketGames Then
        .Range("B2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Game]:[P1 Game]],MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        ",SUMIFS(MatchRecords[[P2 Game]:[P2 Game]],MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)),"""")"
        
        .Range("C2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Game]:[P1 Game]],MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        ",SUMIFS(MatchRecords[[P2 Game]:[P2 Game]],MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)),"""")"
    ElseIf summaryType = bracketSets Then
        .Range("B2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Set]:[P1 Set]],MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        ",SUMIFS(MatchRecords[[P2 Set]:[P2 Set]],MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)),"""")"
        
        .Range("C2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Set]:[P1 Set]],MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)" & _
        ",SUMIFS(MatchRecords[[P2 Set]:[P2 Set]],MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[Bracket]:[Bracket]],$A$1)),"""")"
    ElseIf summaryType = summaryGames Then
        .Range("B2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Game]:[P1 Game]],MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        ",SUMIFS(MatchRecords[[P2 Game]:[P2 Game]],MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[P1 Name]:[P1 Name]],B$1)),"""")"
        
        .Range("C2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Game]:[P1 Game]],MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        ",SUMIFS(MatchRecords[[P2 Game]:[P2 Game]],MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[P1 Name]:[P1 Name]],$A2)),"""")"
    ElseIf summaryType = summarySets Then
        .Range("B2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Set]:[P1 Set]],MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        ",SUMIFS(MatchRecords[[P2 Set]:[P2 Set]],MatchRecords[[P2 Name]:[P2 Name]],$A2,MatchRecords[[P1 Name]:[P1 Name]],B$1)),"""")"

        .Range("C2").Value = "=IF(COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],$A2,MatchRecords[[P2 Name]:[P2 Name]],B$1)" & _
        "+COUNTIFS(MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        "<>0,SUM(SUMIFS(MatchRecords[[P1 Set]:[P1 Set]],MatchRecords[[P1 Name]:[P1 Name]],B$1,MatchRecords[[P2 Name]:[P2 Name]],$A2)" & _
        ",SUMIFS(MatchRecords[[P2 Set]:[P2 Set]],MatchRecords[[P2 Name]:[P2 Name]],B$1,MatchRecords[[P1 Name]:[P1 Name]],$A2)),"""")"
    End If
End With
End Sub
Private Sub FormatLeftColumn()
    With brktSht.Range("A2")
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeRight).Color = -2500135
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeBottom).Color = -16777216
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
End Sub
Private Sub MergeTopCell()
'Purpose: To create Merged cells on topmost row for display reasons.
'Last Updated: July 2021

'Declare variable
    Dim i As Integer
'Merge every other cell in the worksheet, with the max range dependent on the bracket player count.
With brktSht
    For i = 2 To (playerCount * 2) Step 2
        brktSht.Range(.Cells(1, i), .Cells(1, i + 1)).Merge
    Next i
End With
FormatTopRow
End Sub
Private Sub AutofillFormats()
'Autofill top row/column to match bracket player count
    With brktSht
        .Range("A2").AutoFill Destination:=.Range("A2:A" & playerCount + 1), Type:=xlFillFormats
        .Columns(1).AutoFit
        .Range("B2:C2").AutoFill Destination:=.Range("B2:" & maxColumn & "2")
        .Range("B2:" & maxColumn & "2").AutoFill Destination:=.Range("B2:" & maxColumn & playerCount + 1)
    End With
End Sub
Private Sub AddNames()
'Purpose: To add all relevant plajer names
Dim i As Integer, j As Integer
With brktSht
    'Add names to left column
    For i = 2 To playerCount + 1
        .Cells(i, 1).Value = playerDict.Items()(i - 2)
    Next i
    'Add names to upper row
    j = 0
    For i = 2 To (playerCount * 2) Step 2
        .Cells(1, i).Value = playerDict.Items()(j)
        j = j + 1
    Next i
    .Columns("B:" & maxColumn).AutoFit
End With
End Sub
Private Sub RangeA1(bracketName As String)
'Purpose: To make cell A1 match the bracket name, for the sake of the score formulas.
    With brktSht.Range("A1")
        .Value = bracketName
        .HorizontalAlignment = xlHAlignCenter
        .Font.name = "Arial"
        .Font.Size = 10
    End With
End Sub
Private Sub AddNAValues()
'Purpose: To add NA to each set of cells in which a player is matched against himself.
Dim i As Integer, j As Integer
With brktSht
    For i = 2 To playerCount + 1
        For j = 2 To ((playerCount * 2) + 1)
            With .Cells(i, j)
                .Value = "N"
                .HorizontalAlignment = xlHAlignRight
                .Interior.ThemeColor = xlThemeColorLight1
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
            End With
            j = j + 1
            With .Cells(i, j)
                .Value = "A"
                .HorizontalAlignment = xlHAlignLeft
                .Interior.ThemeColor = xlThemeColorLight1
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
            End With
            i = i + 1
        Next j
    Next i
End With
End Sub
Private Sub RemoveBlankFormulas()
'Purpose: To remove all formulas with a blank value.  Done to save on performance.
Dim x As Integer, y As Integer
With brktSht
    For x = 2 To playerCount + 1
        For y = 2 To ((playerCount + 1) * 2)
            If .Cells(x, y) = "" Then
                .Cells(x, y).ClearContents
            End If
        Next y
    Next x
End With
End Sub
Private Sub FinalColumns()
Dim winCol As Range, lossCol As Range, totalCol As Range, scoreCol As Range
With brktSht
'Create Win column
    Set winCol = .Range(maxColumn & "1").Offset(0, 1)
    winCol.Value = "Wins"
    winCol.Offset(1, 0).Value = _
        "=SUMPRODUCT(--(MOD(COLUMN(A2:" & maxColumn & "2)-COLUMN(A2)+1,2)=0),A2:" & maxColumn & "2)"
    winCol.Offset(1, 0).AutoFill _
        Destination:=.Range(winCol.Offset(1, 0), winCol.Offset(playerCount, 0))
        
    With Range(winCol, winCol.Offset(playerCount, 0))
        .HorizontalAlignment = xlHAlignCenter
        .Interior.Color = 65280
        With .Borders(xlEdgeRight)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .Color = -16777216
            .Weight = xlThin
        End With
    End With
'Create loss column
    Set lossCol = .Range(maxColumn & "1").Offset(0, 2)
    lossCol.Value = "Losses"
    lossCol.Offset(1, 0).Value = _
        "=SUMPRODUCT(--(MOD(COLUMN(B2:" & maxColumn & "2)-COLUMN(B2)+1,2)=0),B2:" & maxColumn & "2)"
    lossCol.Offset(1, 0).AutoFill _
        Destination:=Range(lossCol.Offset(1, 0), lossCol.Offset(playerCount, 0))
        
    With Range(lossCol, lossCol.Offset(playerCount, 0))
        .HorizontalAlignment = xlHAlignCenter
        .Interior.Color = 39423
        With .Borders(xlEdgeRight)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .Color = -16777216
            .Weight = xlThin
        End With
    End With
'Create Sets/Games column
    Set totalCol = .Range(maxColumn & "1").Offset(0, 3)
    totalCol.Value = "Total"
    totalCol.Offset(1, 0).FormulaR1C1 = "=SUM(RC[-2],RC[-1])"
    totalCol.Offset(1, 0).AutoFill _
        Destination:=Range(totalCol.Offset(1, 0), totalCol.Offset(playerCount, 0))

    With Range(totalCol, totalCol.Offset(playerCount, 0))
        .HorizontalAlignment = xlHAlignCenter
        .Interior.Color = 3326705
        With .Borders(xlEdgeRight)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .Color = -16777216
            .Weight = xlThin
        End With
    End With
'Create Score column
    Set scoreCol = .Range(maxColumn & "1").Offset(0, 4)
    scoreCol.Value = "Score"
    scoreCol.Offset(1, 0).FormulaR1C1 = "=TRUNC((RC[-3]/RC[-1]),2)"
    scoreCol.Offset(1, 0).AutoFill _
        Destination:=Range(scoreCol.Offset(1, 0), scoreCol.Offset(playerCount, 0))
        
    With Range(scoreCol, scoreCol.Offset(playerCount, 0))
        .HorizontalAlignment = xlHAlignCenter
        .Interior.Color = 65535
        With .Borders(xlEdgeRight)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .Color = -16777216
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .Color = -16777216
            .Weight = xlThin
        End With
    End With
End With
End Sub
Private Sub FreezePanes()
Dim wb As Workbook
Set wb = brktSht.Parent
    With wb.Windows(1)
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
'Format cell borders of topmost row/column.
Private Sub FormatTopRow()
    With brktSht.Range("B1:" & maxColumn & "1")
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeRight).Color = -2500135
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = -2500135
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeLeft).Color = -2500135
        .Borders(xlInsideVertical).Weight = xlThick
        .Borders(xlInsideVertical).Color = -2500135
    End With
End Sub
Private Function WorksheetExists(wb As Workbook, shtName As String) As Boolean
'Purpose: To check for any worksheet name that already exists in the workbook.
    'There is a public function for this, but having it in the class module makes it easier to keep track of.
'Last Updated: September 2021

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
