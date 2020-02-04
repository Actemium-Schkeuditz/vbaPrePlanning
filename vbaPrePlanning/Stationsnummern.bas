Attribute VB_Name = "Stationsnummern"
' Skript zur Ermittlung der Stationsnummern der IO-Racks
' V0.1
' nicht fertig
' 24.01.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de


Option Explicit

Public Sub RACK_STATIONSNUMMERN()


    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim y As Integer
    Dim spalteStationsNummer As String
    Dim spalteKWS_StationsNummer As String
    'Dim spalteEinbauortRack As String
    
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets(tabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

   
        spalteKWS_StationsNummer = "BC"
        spalteStationsNummer = "BU"
        'spalteEinbauortRack = "BV"
   
    
        'Umkopieren
        ' Daten schreiben
        For i = 3 To zeilenanzahl
            ' Prüfe ob Stationsnummer mit Eintrag
            If Cells(i, spalteKWS_StationsNummer) <> "" Then
                If IsNumeric(Cells(i, spalteKWS_StationsNummer)) Then
                    ' Mache
                    Cells(i, spalteStationsNummer) = Cells(i, spalteKWS_StationsNummer)
                Else
                    Cells(i, spalteStationsNummer) = Cells(i, spalteKWS_StationsNummer)
                    Cells(i, spalteStationsNummer).Interior.ColorIndex = 3
                    MsgBox "Stationsnummer Prüfen!  Zeile: " + Str(i)
                End If
            End If
    
        Next i
    
    
    End With

End Sub

