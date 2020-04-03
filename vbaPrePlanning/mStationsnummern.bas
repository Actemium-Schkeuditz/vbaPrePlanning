Attribute VB_Name = "mStationsnummern"
' Skript zur Ermittlung der Stationsnummern der IO-Racks
' V0.3
' nicht fertig
' 03.04.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Stationsnummern)

Option Explicit

Public Sub RACK_STATIONSNUMMERN()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim TabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim iKanal As Long
    Dim SpalteStationsnummer As String
    Dim spalteKWS_StationsNummer As String
    Dim sSpalteStationsnummerSignal As String
    Dim iSpalteStationsnummerSignal As Long
    Dim ExcelConfig As New cExcelConfig
      
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten
   
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](TabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl
 
        spalteKWS_StationsNummer = ExcelConfig.StationsnummerKWS
        SpalteStationsnummer = ExcelConfig.Stationsnummer
        sSpalteStationsnummerSignal = ExcelConfig.StationsnummerSignal_1
        
        iSpalteStationsnummerSignal = SpaltenBuchstaben2Int(sSpalteStationsnummerSignal)
    
        'Umkopieren
        ' Daten schreiben
        For i = 3 To zeilenanzahl
            ' Prüfe ob Stationsnummer mit Eintrag
            If .Cells.Item(i, spalteKWS_StationsNummer) <> vbNullString Then
                If IsNumeric(.Cells.Item(i, spalteKWS_StationsNummer)) Then
                    ' Mache
                    .Cells.Item(i, SpalteStationsnummer) = .Cells.Item(i, spalteKWS_StationsNummer)
                Else
                    .Cells.Item(i, SpalteStationsnummer) = .Cells.Item(i, spalteKWS_StationsNummer)
                    .Cells.Item(i, SpalteStationsnummer).Interior.ColorIndex = 3
                    MsgBox "Stationsnummer Prüfen!  Zeile: " + str(i)
                End If
                 For iKanal = 1 To 6
                    If .Cells.Item(i, iSpalteStationsnummerSignal + 2 + (14 * (iKanal - 1))) <> vbNullString Then ' prüfen ob Typ-Station nicht leer
                        If IsNumeric(.Cells.Item(i, spalteKWS_StationsNummer)) Then
                            ' Mache
                            .Cells.Item(i, iSpalteStationsnummerSignal + (14 * (iKanal - 1))) = .Cells.Item(i, spalteKWS_StationsNummer)
                        Else
                            .Cells.Item(i, iSpalteStationsnummerSignal + (14 * (iKanal - 1))) = .Cells.Item(i, spalteKWS_StationsNummer)
                            .Cells.Item(i, iSpalteStationsnummerSignal + (14 * (iKanal - 1))).Interior.ColorIndex = 3
                            MsgBox "Stationsnummer Prüfen!  Zeile: " + str(i)
                        End If
                    End If
                Next iKanal
            End If
        Next i
    End With
End Sub
