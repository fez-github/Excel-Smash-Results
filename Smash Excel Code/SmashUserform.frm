VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SmashUserform 
   Caption         =   "UserForm1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   OleObjectBlob   =   "SmashUserform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SmashUserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_AllBracketSheets_Click()
    Dim brktSht As New BracketSheet
    For i = 0 To SmashUserform.Bracket_ComboBox.ListCount - 1
        If Bracket_ComboBox.List(i) = "" Then '
            MsgBox "There are no brackets to add."
            Exit Sub
        End If
        If WorksheetExists(ThisWorkbook, SmashUserform.Bracket_ComboBox.List(i) & "(G)") = False Then
            brktSht.InitiateModule SmashUserform.Bracket_ComboBox.List(i), bracketGames
            brktSht.InitiateModule SmashUserform.Bracket_ComboBox.List(i), bracketSets
        End If
    Next i
'    For i = 1 To Range("BracketTable").Rows.Count
'        brktSht.InitiateModule Range("BracketTable[Bracket Name]")(i).Value, bracketGames
'        brktSht.InitiateModule Range("BracketTable[Bracket Name]")(i).Value, bracketSets
'    Next i
    SheetNamePopulate
End Sub

Private Sub btn_NewBracketSheet_Click()
    If IsEmpty(Bracket_ComboBox.Value) = True Or Bracket_ComboBox.Value = "" Then
        MsgBox "Please enter a valid bracket name!"
        Exit Sub
    End If
    If WorksheetExists(ThisWorkbook, Bracket_ComboBox.Value) = True Then
        MsgBox "That sheet name already exists.  Please enter a new sheet name."
        Exit Sub
    End If
    Dim brktSht As New BracketSheet
    brktSht.InitiateModule Bracket_ComboBox.Value, bracketGames
    brktSht.InitiateModule Bracket_ComboBox.Value, bracketSets
    SheetNamePopulate
End Sub
Private Sub btn_SheetNavigation_Click()
    Sheets(SmashUserform.ComboBox1.Value).Activate
End Sub
Private Sub btn_SummarizeBrackets_Click()
    Dim brktSht As New BracketSheet
    If WorksheetExists(ThisWorkbook, "AllBrackets(G)") Or WorksheetExists(ThisWorkbook, "AllBrackets(S)") Then
        MsgBox "Summary sheets have already been made.  There is no need for extras."
        Exit Sub
    End If
    If Bracket_ComboBox.Value = "" Then
        MsgBox "There are no brackets to summarize."
        Exit Sub
    End If
    brktSht.InitiateModule "", summaryGames
    brktSht.InitiateModule "", summarySets
    SheetNamePopulate
End Sub

Private Sub cmd_ObtainBracket_Click()
    If txt_ChallongeUser.Value = "" Then
        MsgBox "No value in Challonge user!  Please insert your Challonge username."
        Exit Sub
    End If
    If txt_BracketName.Value = "" Then
        MsgBox "No value in Bracket name!  Please insert the bracket ID found in the Challonge URL."
        Exit Sub
    End If
    If txt_APIKey.Value = "" Then
        MsgBox "No value in API Key!  Please insert your API key.  This can be found in your Challonge settings."
        Exit Sub
    End If
    ObtainBracketAndMatches SmashUserform.txt_BracketName, SmashUserform.txt_ChallongeUser, SmashUserform.txt_APIKey
    SheetNamePopulate
    BracketNamePopulate
End Sub
Private Sub CommandButton3_Click()
    Sheets(SmashUserform.ComboBox1.Value).Activate
End Sub
Private Sub CommandButton4_Click()
    If IsEmpty(TextBox1.Value) = True Then
        MsgBox "Please enter a file name."
        Exit Sub
    Else
        Export TextBox1.Value
    End If
End Sub
Private Sub CommandButton5_Click()

End Sub
Public Sub UserForm_Initialize()
Dim i As Integer
'Populate Brackets ComboBox.
    BracketNamePopulate
'Populate sheets ComboBox.
    SheetNamePopulate
'Load default Export name
    SmashUserform.TextBox1.Value = Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5)
End Sub
Private Sub BracketNamePopulate()
Dim allBrackets As Object
    With SmashUserform.Bracket_ComboBox
        If WorksheetExists(ThisWorkbook, "Match Records") = True Then
            .Clear
            'Public function to add all unique bracket names.
                Set allBrackets = Unique_Values(allBrackets, Range("MatchRecords").Cells(2, 1))
            For Each varKey In allBrackets.Keys()
                .AddItem allBrackets(varKey)
            Next
'            For i = 1 To Range("MatchRecords").Rows.count
'                If IsEmpty(Range("MatchRecords[Bracket]")(i)) = False Then
'                    .AddItem Range("MatchRecords[Bracket]")(i).value
'                End If
'            Next i
        End If
        .ListIndex = 0
    End With
End Sub
Private Sub SheetNamePopulate()
    With SmashUserform.ComboBox1
        .Clear
        For Each sht In ThisWorkbook.Sheets
            If sht.name <> "Userform" Then
                .AddItem sht.name
                .ListIndex = 0
            End If
        Next sht
    End With
End Sub
