Attribute VB_Name = "ImportChallonge"
Sub QuickTest()
'Purpose: To easily test this code without interacting with the Userform.
    'ObtainBracketAndMatches bracket URL, Username, API key
End Sub
Sub ObtainBracketAndMatches(url As String, userName As String, key As String)
'Purpose: Performs initial API call, and then handles generation/insertion of tables and records.
'Last Updated: September 27, 2021

'Declare variables
    Dim APItournament As Object, APIparticipants As Object, APImatches As Object
    Dim i As Integer, j As Integer
    Dim resultsTable As ListObject, bracketTable As ListObject, placingsTable As ListObject
    Dim brktName As String, player1 As String, player2 As String
    Dim placingsColumn As Range
    Dim dataImport As New DataImporter  'Class Module
'Acquire API information.  Parsed with JsonConverter module.
    Set APItournament = ParseJson(API_Get("GET", "https://api.challonge.com/v1/tournaments/" & url & ".json", userName, key))
    Set APIparticipants = ParseJson(API_Get("GET", "https://api.challonge.com/v1/tournaments/" & url & "/participants.json", userName, key))
    Set APImatches = ParseJson(API_Get("GET", "https://api.challonge.com/v1/tournaments/" & url & "/matches.json", userName, key))
    brktName = APItournament("tournament")("name")
'Add bracket name to table
    Set bracketTable = SetTable("Brackets", "BracketTable")
    If bracketTable Is Nothing Then
        Set bracketTable = dataImport.CreateBracketTable(Sheets("Brackets"))
    End If
'Add bracket information to table.
        dataImport.UpdateBracketTable bracketTable, _
            APItournament("tournament")("name"), _
            Left(APItournament("tournament")("created_at"), 10), _
            "http://www.challonge.com/" & APItournament("tournament")("url"), _
            APItournament("tournament")("participants_count"), _
            StrConv(APItournament("tournament")("tournament_type"), vbProperCase)
'Add new participants to list, and their placings.
    Set placingsTable = SetTable("Placings", "PlacingsTable")
    If placingsTable Is Nothing Then
        Set placingsTable = dataImport.CreatePlacingsTable(Sheets("Placings"))
    End If
    'Set column value for next function.
        Set placingsColumn = dataImport.GetBracketColumn(placingsTable, brktName)
    'For each participant, add their name and rank of the bracket.
        For i = 1 To APIparticipants.Count
            dataImport.AddPlacingsRank placingsTable, placingsColumn, _
                APIparticipants(i)("participant")("name"), _
                APIparticipants(i)("participant")("final_rank"), _
                APIparticipants.Count
        Next i
    'Add NA values to empty cells of the table.
        dataImport.AddPlacingsNA placingsTable
'Search for match records table, and declare as variable.  If not, make new one.
    Set resultsTable = SetTable("Match Records", "MatchRecords")
    If resultsTable Is Nothing Then
        Set resultsTable = dataImport.CreateMatchRecords(Sheets("Match Records"))
    End If
'Add match results
    
    For i = 1 To APImatches.Count
        'Suggested_play_order has been returning null.  This check is used to take that into account.
        'Find player 1 name.  Matches API only includes id.
        For j = 1 To APIparticipants.Count
            If APImatches(i)("match")("player1_id") = APIparticipants(j)("participant")("id") Then
                player1 = APIparticipants(j)("participant")("name")
                Exit For
            End If
        Next j
            'Repeat loop for player 2.
        For j = 1 To APIparticipants.Count
            If APImatches(i)("match")("player2_id") = APIparticipants(j)("participant")("id") Then
                player2 = APIparticipants(j)("participant")("name")
                Exit For
            End If
        Next j
        dataImport.AddMatchRecords resultsTable, _
                brktName, _
                Left(APImatches(i)("match")("started_at"), 10), _
                APImatches(i)("match")("identifier"), _
                player1, _
                Left(APImatches(i)("match")("scores_csv"), 1), _
                player2, _
                Right(APImatches(i)("match")("scores_csv"), 1)
                'suggested_play_order is returning null right now.
    Next i
    
    dataImport.SortMatchRecords resultsTable
MsgBox "All data has been imported!"
End Sub
Function SetTable(shtName As String, tblName As String) As ListObject
'Purpose: To search for the proper table and set the listobject to it.  If no table is found, create one.
'Last Updated: September 27, 2021
    Dim tbl As ListObject
    If WorksheetExists(ThisWorkbook, shtName) = False Then
        ThisWorkbook.Sheets.Add.name = shtName
    End If
    For Each tbl In Sheets(shtName).ListObjects
        If tbl.name = tblName Then
            Set SetTable = ThisWorkbook.Sheets(shtName).ListObjects(tblName)
            Exit Function
        End If
    Next
End Function
