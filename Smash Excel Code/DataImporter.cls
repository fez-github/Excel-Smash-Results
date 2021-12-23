VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function CreateBracketTable(sht As Worksheet) As ListObject
'Purpose: To create a new Bracket table if one doesn't already exist.
'Last Updated: September 27, 2021
'NOTE: I normally use a Custom Table Style for brackets.  Look into making one later.

'Declare variables
    Dim tbl As ListObject
    Set tbl = sht.ListObjects.Add(Source:=sht.Range("A1:D2"), XlListObjectHasHeaders:=1, Destination:=sht.Range("A1"), TableStyleName:="TableStyleLight8")
'Create headers
    tbl.HeaderRowRange(, 1).Value = "Bracket Name"
    tbl.HeaderRowRange(, 2).Value = "Bracket Date"
    tbl.HeaderRowRange(, 3).Value = "Bracket Link"
    tbl.HeaderRowRange(, 4).Value = "# of Players"
    tbl.HeaderRowRange(, 5).Value = "Format"
    tbl.name = "BracketTable"
    Set CreateBracketTable = tbl
End Function
Sub UpdateBracketTable(tbl As ListObject, name As Variant, brktDate As Variant, url As Variant, playerCount As Variant, brktFormat As Variant)
'Purpose: to add any new brackets to the general bracket table.
'Last Updated: September 27, 2021

With tbl
    'Add headers
        .ListRows.Add
        .DataBodyRange(.ListRows.Count, 1).Value = name
        .DataBodyRange(.ListRows.Count, 2).Value = brktDate
        .DataBodyRange(.ListRows.Count, 3).Value = url
        .DataBodyRange(.ListRows.Count, 4).Value = playerCount
        .DataBodyRange(.ListRows.Count, 5).Value = brktFormat
    'Remove any excess rows.
        If IsEmpty(.DataBodyRange(1, 1)) = True Then
            .ListRows(1).Delete
        End If
    'Sort table by Bracket name.
        .Sort.SortFields.Clear
        .Sort.SortFields.Add key:=.HeaderRowRange(, 1), Order:=xlAscending
    'Remove any duplicates.
        .DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlYes
End With
End Sub
Function CreatePlacingsTable(sht As Worksheet)
'Purpose: To create a new placings table in case if there isn't one already.
'Last Updated: September 27, 2021

'Declare variable
    Dim tbl As ListObject
    Set tbl = sht.ListObjects.Add(Source:=sht.Range("A1:C2"), XlListObjectHasHeaders:=1, Destination:=sht.Range("A1"), TableStyleName:="TableStyleLight8")
    tbl.name = "PlacingsTable"
'Add Player, empty space, and attendance as headers.
    tbl.HeaderRowRange(, 1).Value = "Player Name"
        tbl.DataBodyRange(1, 1).Font.Bold = True
    tbl.HeaderRowRange(, 2).Value = "Bracket 1"
    tbl.HeaderRowRange(, 3).Value = "Attendance"
        tbl.DataBodyRange(1, 1).HorizontalAlignment = xlHAlignCenter
        tbl.DataBodyRange(1, 2).HorizontalAlignment = xlHAlignCenter
        tbl.DataBodyRange(1, 3).HorizontalAlignment = xlHAlignCenter
'Add new row, with Attendance column calculated.
    'tbl.DataBodyRange(1, 3).Value = _
        "=IF((COUNTIFS(A2:B2,""<>NA"",A2:B2,""<>""&"""")-1)=-1,0,(COUNTIFS(A2:B2,""<>NA"",A2:B2,""<>""&"""")-1))"
    tbl.DataBodyRange(1, 3).Value = "=IF((COUNTIFS(OFFSET(PlacingsTable[[#Headers],[Player Name]],ROW([@Player Name])-1,,,COLUMN(PlacingsTable[[#Headers],[Attendance]])-1),""<>NA"",OFFSET(PlacingsTable[[#Headers],[Player Name]],ROW([@Player Name])-1,,,COLUMN(PlacingsTable[[#Headers],[Attendance]])-1),""<>""&"""")-1)=-1,0,(COUNTIFS(OFFSET(PlacingsTable[[#Headers],[Player Name]],ROW([@Player Name])-1,,,COLUMN(PlacingsTable[[#Headers]" & _
        ",[Attendance]])-1),""<>NA"",OFFSET(PlacingsTable[[#Headers],[Player Name]],ROW([@Player Name])-1,,,COLUMN(PlacingsTable[[#Headers],[Attendance]])-1),""<>""&"""")-1))" & _
        ""  'My apologies.  :(
'Generate conditional formatting for Attendance column.
With tbl.DataBodyRange(1, 3)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="3"
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="3"
    With .FormatConditions(1)   'Attendance less than 3
        .Interior.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
    With .FormatConditions(2)   'Attendance at least 3
        .Interior.Color = RGB(146, 208, 80)
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
End With
'Generate conditional formatting for placement columns.
With tbl.DataBodyRange(1, 2)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA"
    With .FormatConditions(1)   'NA value is placed in bracket column.
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End With
'Generate conditional formatting for player names.(duplicate check)
With tbl.DataBodyRange(1, 1)
    .FormatConditions.AddUniqueValues
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    .FormatConditions(1).DupeUnique = xlDuplicate
    With .FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With .FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
End With
    Set CreatePlacingsTable = tbl
End Function
Function GetBracketColumn(tbl As ListObject, bracketName As String) As Range
'Purpose: To identify the bracket column to perform a loop on.
'Last Updated: December 3, 2021
With tbl
'Check if there are duplicate column names.  If so, delete.
    If .HeaderRowRange.Find(What:=bracketName) Is Nothing Then
        'Must check if this is needed, as you won't need to add one for the first bracket.
            If IsEmpty(.DataBodyRange(2, .ListColumns.Count).Offset(, -1)) = False Then
                .ListColumns.Add (.ListColumns.Count)
            End If
        .HeaderRowRange(, .ListColumns.Count - 1).Value = bracketName
        Set GetBracketColumn = .HeaderRowRange(, .ListColumns.Count - 1)
    Else
        Set GetBracketColumn = .HeaderRowRange.Find(What:=bracketName)
    End If
End With
End Function
Sub AddPlacingsRank(tbl As ListObject, brktCol As Range, name, rank, playerCount As Integer)
'Purpose: To add all placings values from the API call into the Placings table.

'Declare variables
    Dim plyrRng As Range
With tbl
'Loop through dictionary.  If player already exists, add ranking in bracket cell.  Otherwise, add to bottom of row, then add new value.
    For i = 1 To playerCount
        Set plyrRng = Sheets("Placings").Range("PlacingsTable[Player Name]").Find(What:=name)
        If plyrRng Is Nothing Then
            .ListRows.Add
            .DataBodyRange(.ListRows.Count, 1).Value = name
            .DataBodyRange(.ListRows.Count, brktCol.Column).Value = _
                rank & " of " & playerCount
        Else
            .DataBodyRange(plyrRng.Row - 1, brktCol.Column).Value = _
                rank & " of " & playerCount
        End If
    Next i
End With
End Sub
Sub AddPlacingsNA(tbl As ListObject)
'Purpose: To run through all bracket cells, and add NA to any blank cells.

'Declare variables
    Dim i As Integer, j As Integer
With tbl
    'Loop through all cells and add NA to blank cells.
        For i = 1 To .ListRows.Count
            For j = 2 To (.ListColumns.Count - 1)
                If IsEmpty(.DataBodyRange(i, j)) = True Then
                    .DataBodyRange(i, j).Value = "NA"
                End If
            Next j
        Next i
    'Delete any excess rows.
        If IsEmpty(.DataBodyRange(1, 1)) = True Then
            .ListRows(1).Delete
        End If
End With
End Sub
Private Sub UpdatePlacingsTable(tbl As ListObject, bracketName As String, placingsAPI As Object)
'-----NOT NEEDED ANYMORE-----
'Purpose: To add new players and/or placements to the placings table.
'Last Updated: September 27, 2021

'Declare variables
    Dim i As Integer, j As Integer
    Dim plyrRng As Range, brktCol As Range
With tbl
    'Check if there are duplicate column names.  If so, delete.
        If .HeaderRowRange.Find(What:=bracketName) Is Nothing Then
            'Must check if this is needed, as you won't need to add one for the first bracket.
                If IsEmpty(.DataBodyRange(2, .ListColumns.Count).Offset(, -1)) = False Then
                    .ListColumns.Add (.ListColumns.Count)
                End If
            .HeaderRowRange(, .ListColumns.Count - 1).Value = bracketName
            Set brktCol = .HeaderRowRange(, .ListColumns.Count - 1)
        Else
            Set brktCol = .HeaderRowRange.Find(What:=bracketName)
        End If
    'Loop through dictionary.  If player already exists, add ranking in bracket cell.  Otherwise, add to bottom of row, then add new value.
        For i = 1 To placingsAPI.Count
            Set plyrRng = Sheets("Placings").Range("PlacingsTable[Player Name]").Find(What:=placingsAPI(i)("participant")("name"))
            If plyrRng Is Nothing Then
                .ListRows.Add
                .DataBodyRange(.ListRows.Count, 1).Value = placingsAPI(i)("participant")("name")
                .DataBodyRange(.ListRows.Count, brktCol.Column).Value = _
                    placingsAPI(i)("participant")("final_rank") & " of " & placingsAPI.Count
            Else
                .DataBodyRange(plyrRng.Row - 1, brktCol.Column).Value = _
                    placingsAPI(i)("participant")("final_rank") & " of " & placingsAPI.Count
            End If
        Next i
    'Once dictionary is iterated, run through each cell that did not have a player present, and write NA in bracket cell.
        For i = 1 To .ListRows.Count
            For j = 2 To (.ListColumns.Count - 1)
                If IsEmpty(.DataBodyRange(i, j)) = True Then
                    .DataBodyRange(i, j).Value = "NA"
                End If
            Next j
        Next i
    'Delete any excess rows.
        If IsEmpty(.DataBodyRange(1, 1)) = True Then
            .ListRows(1).Delete
        End If
End With
End Sub
Function CreateMatchRecords(sht As Worksheet)
'Purpose: Creates a new Match Records table if one doesn't exist.
'Last Updated: September 27, 2021
'Declare variables
    Dim tbl As ListObject
    Set tbl = sht.ListObjects.Add(Source:=sht.Range("A1:J2"), XlListObjectHasHeaders:=1, Destination:=sht.Range("A1"), TableStyleName:="TableStyleLight8")
'Set headers
    tbl.HeaderRowRange(, 1).Value = "Bracket"
    tbl.HeaderRowRange(, 2).Value = "Date"
    tbl.HeaderRowRange(, 3).Value = "Round #"
    tbl.HeaderRowRange(, 4).Value = "P1 Name"
    tbl.HeaderRowRange(, 5).Value = "P1 Game"
    tbl.HeaderRowRange(, 6).Value = "P1 Set"
    tbl.HeaderRowRange(, 6).EntireColumn.Hidden = True
    tbl.HeaderRowRange(, 7).Value = "P2 Name"
    tbl.HeaderRowRange(, 8).Value = "P2 Game"
    tbl.HeaderRowRange(, 9).Value = "P2 Set"
    tbl.HeaderRowRange(, 9).EntireColumn.Hidden = True
    tbl.HeaderRowRange(, 10).Value = "Win?"
    tbl.HeaderRowRange(, 10).EntireColumn.Hidden = True
'Add formulas for hidden columns.  These are not to be touched.
    tbl.HeaderRowRange(2, 6).FormulaR1C1 = "=IF([[P1 Game]]>[[P2 Game]],1,0)"
    tbl.HeaderRowRange(2, 9).FormulaR1C1 = "=IF([[P2 Game]]>[[P1 Game]],1,0)"
    tbl.HeaderRowRange(2, 10).FormulaR1C1 = "=IF([[P1 Game]]>[[P2 Game]],""Win"",""Loss"")"
    tbl.name = "MatchRecords"
    Set CreateMatchRecords = tbl
End Function
Sub AddMatchRecords(tbl As ListObject, bracketName As String, bracketDate, roundNum, p1Name As String, p1Score As String, p2Name As String, p2Score As String)
'Purpose: Loop
With tbl
     .ListRows.Add
     'Bracket Name
         .DataBodyRange(.ListRows.Count, 1).Value = bracketName
     'Bracket Date
         .DataBodyRange(.ListRows.Count, 2).Value = bracketDate
     'Round Number
         .DataBodyRange(.ListRows.Count, 3).Value = roundNum
     'Player 1 Name
        .DataBodyRange(.ListRows.Count, 4).Value = p1Name
     'Player 1 Score
         If p1Score > 3 Or p1Score < 0 Or IsNumeric(p1Score) = False Then   'Check for invalid values due to DQ.
             .DataBodyRange(.ListRows.Count, 5).Value = 0
         Else
             .DataBodyRange(.ListRows.Count, 5).Value = p1Score
         End If
     'Player 2 Name
        .DataBodyRange(.ListRows.Count, 7).Value = p2Name
     'Player 2 Score
         If p2Score > 3 Or p2Score < 0 Or IsNumeric(p2Score) = False Then   'Check for invalid values due to DQ.
             .DataBodyRange(.ListRows.Count, 8).Value = 0
         Else
             .DataBodyRange(.ListRows.Count, 8).Value = p2Score
         End If
End With
End Sub
Sub SortMatchRecords(tbl As ListObject)
With tbl
    'Sort all records by bracket name, then by round number.
        .Sort.SortFields.Clear
        .Sort.SortFields.Add key:=.HeaderRowRange(, 1), Order:=xlAscending
        .Sort.SortFields.Add key:=.HeaderRowRange(, 3), Order:=xlAscending
        .Sort.Apply
    'Remove any potential blank rows.
        If IsEmpty(.DataBodyRange(1, 1)) = True Then
            .ListRows(1).Delete
        End If
    'Check for duplicates and delete them.
        .DataBodyRange.RemoveDuplicates Columns:=Array(1, 3), Header:=xlYes
End With
End Sub
Private Sub UpdateMatchRecords(tbl As ListObject, bracketName As String, playersAPI As Object, matchesAPI As Object)
'Purpose: To add new rows to the Match Records field.
'Last Updated: September 27, 2021

'Declare variable for looping
    Dim i As Integer
'Add each new record to the table.
With tbl
    For i = 1 To matchesAPI.Count
        .ListRows.Add
        'Bracket Name
            .DataBodyRange(.ListRows.Count, 1).Value = bracketName
        'Bracket Date
            .DataBodyRange(.ListRows.Count, 2).Value = Left(matchesAPI(i)("match")("started_at"), 10)
        'Round Number
            .DataBodyRange(.ListRows.Count, 3).Value = matchesAPI(i)("match")("suggested_play_order")
        'Round indicator(Winner's Round 1, Loser's Round 2, etc.)
            'Cells(i, 4).value = matchesAPI(matchesAPI.Count, matchesAPI.Count - 1, matchesAPI(i)("match")("round"))
            'cannot be done now.  Must know whether or not True Grands happens.)
       
        'Player 1 Name
            For j = 1 To playersAPI.Count
                If matchesAPI(i)("match")("player1_id") = playersAPI(j)("participant")("id") Then
                    .DataBodyRange(.ListRows.Count, 4).Value = playersAPI(j)("participant")("name")
                    Exit For
                End If
            Next j
        'Player 1 Score
            If Left(matchesAPI(i)("match")("scores_csv"), 1) > 3 Or Left(matchesAPI(i)("match")("scores_csv"), 1) < 0 Or _
            IsNumeric(Left(matchesAPI(i)("match")("scores_csv"), 1)) = False Then   'Check for invalid values due to DQ.
                .DataBodyRange(.ListRows.Count, 5).Value = 0
            Else
                .DataBodyRange(.ListRows.Count, 5).Value = Left(matchesAPI(i)("match")("scores_csv"), 1)
            End If
        'Player 2 Name
            For j = 1 To playersAPI.Count
                If matchesAPI(i)("match")("player2_id") = playersAPI(j)("participant")("id") Then
                    .DataBodyRange(.ListRows.Count, 7).Value = playersAPI(j)("participant")("name")
                    Exit For
                End If
            Next j
        'Player 2 Score
            If Right(matchesAPI(i)("match")("scores_csv"), 1) > 3 Or Right(matchesAPI(i)("match")("scores_csv"), 1) < 0 Or _
            IsNumeric(Right(matchesAPI(i)("match")("scores_csv"), 1)) = False Then    'Check for invalid values due to DQ.
                .DataBodyRange(.ListRows.Count, 8).Value = 0
            Else
                .DataBodyRange(.ListRows.Count, 8).Value = Right(matchesAPI(i)("match")("scores_csv"), 1)
            End If
    Next i
    'Sort all records by bracket name, then by round number.
        .Sort.SortFields.Clear
        .Sort.SortFields.Add key:=.HeaderRowRange(, 1), Order:=xlAscending
        .Sort.SortFields.Add key:=.HeaderRowRange(, 3), Order:=xlAscending
        .Sort.Apply
    'Remove any potential blank rows.
        If IsEmpty(.DataBodyRange(1, 1)) = True Then
            .ListRows(1).Delete
        End If
    'Check for duplicates and delete them.
        .DataBodyRange.RemoveDuplicates Columns:=Array(1, 3), Header:=xlYes
End With
End Sub
Private Sub Class_Initialize()

'This variable is used to handle a workaround for the "suggested_play_order" value of APIMatches returning null.
'Public totalRoundCount As Integer
End Sub
