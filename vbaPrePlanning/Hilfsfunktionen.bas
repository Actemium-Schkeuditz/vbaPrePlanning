Attribute VB_Name = "Hilfsfunktionen"
' Hilfsfunktionen die öfters benötigt werden
' V0.4
' 10.02.2020
' neu: SortTable
' Christian Langrock
' christian.langrock@actemium.de

Option Explicit

Public Function SpaltenBuchstaben2Int(pSpalte As String) As Integer
    'ermittel der Spaltennummer aus den Spaltenbuchstaben
    SpaltenBuchstaben2Int = Columns(pSpalte).Column


End Function

Public Sub SortTable(tablename As String, SortSpalte1 As String, SortSpalte2 As String, Optional SortSpalte3 As String)
    ' sortieren von Daten nach drei oder zwei Spalten
    ' Aufrufen der Tabele und Auswählen dieser
    ThisWorkbook.Worksheets(tablename).Activate
    ' Anzeige ausschalten
    Application.ScreenUpdating = False
        
    Dim rTable As Range
    Set rTable = Range("A3")                     ' Ohne die Überschriften
    If SortSpalte3 = vbNullString Then
        rTable.Sort _
        Key1:=Range(SortSpalte1 & "3"), Order1:=xlAscending, _
        Key2:=Range(SortSpalte2 & "3"), Order1:=xlAscending, _
        Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
    Else
        rTable.Sort _
        Key1:=Range(SortSpalte1 & "3"), Order1:=xlAscending, _
        Key2:=Range(SortSpalte2 & "3"), Order1:=xlAscending, _
        Key3:=Range(SortSpalte3 & "3"), Order1:=xlAscending, _
        Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
    End If
        
BeforeExit:
    Set rTable = Nothing
    Exit Sub
End Sub

Function returnStation(data As cKanalBelegungen) As Collection
    ' collect all Stations without duplicates
     On Error Resume Next
    Set returnStation = Nothing

    Dim col As New Collection
    Dim bSearchinCol As Boolean
    Dim it As Variant
    Dim sData As New cBelegung
    
    For Each sData In data
        ' prüfen ob Stationsnummer schon in Collection
            bSearchinCol = True
        For Each it In col
            If it = sData.Stationsnummer Then
                bSearchinCol = False
                
            End If
        Next
        If bSearchinCol = True Then
            col.Add sData.Stationsnummer         '  dynamically add value to the end
        End If
    Next
    ' Rückgabe der Daten
    Set returnStation = col
End Function


