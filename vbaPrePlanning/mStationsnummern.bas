Attribute VB_Name = "mStationsnummern"
' Skript zur Ermittlung der Stationsnummern der IO-Racks
' V0.1
' nicht fertig
' 24.01.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Stationsnummern)

Option Explicit

Public Sub RACK_STATIONSNUMMERN()


    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    'Dim y As Long
    Dim spalteStationsnummer As String
    Dim spalteKWS_StationsNummer As String
    'Dim spalteEinbauortRack As String
    
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

   
        spalteKWS_StationsNummer = "BC"
        spalteStationsnummer = "BU"
        'spalteEinbauortRack = "BV"
   
    
        'Umkopieren
        ' Daten schreiben
        For i = 3 To zeilenanzahl
            ' Prüfe ob Stationsnummer mit Eintrag
            If .Cells.Item(i, spalteKWS_StationsNummer) <> vbNullString Then
                If IsNumeric(.Cells.Item(i, spalteKWS_StationsNummer)) Then
                    ' Mache
                    .Cells.Item(i, spalteStationsnummer) = .Cells.Item(i, spalteKWS_StationsNummer)
                Else
                    .Cells.Item(i, spalteStationsnummer) = .Cells.Item(i, spalteKWS_StationsNummer)
                    .Cells.Item(i, spalteStationsnummer).Interior.ColorIndex = 3
                    MsgBox "Stationsnummer Prüfen!  Zeile: " + str(i)
                End If
            End If
    
        Next i
    
    
    End With

End Sub

