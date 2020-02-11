Attribute VB_Name = "mHilfsfunktionen"
' Hilfsfunktionen die öfters benötigt werden
' V0.4
' 10.02.2020
' neu: SortTable
' Christian Langrock
' christian.langrock@actemium.de
'@folder Hilfsfunktionen
Option Explicit

Public Function SpaltenBuchstaben2Int(ByRef pSpalte As String) As Long
    'ermittel der Spaltennummer aus den Spaltenbuchstaben
    SpaltenBuchstaben2Int = Columns(pSpalte).Column


End Function

Public Sub SortTable(ByRef tablename As String, ByRef SortSpalte1 As String, ByRef SortSpalte2 As String, Optional ByRef SortSpalte3 As String)
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
        

    Set rTable = Nothing
    Exit Sub
End Sub








