Attribute VB_Name = "mSPSRackBMK"
' Skript zur Ermittlung der Anlagen und Ortskennzeichen der IO-Racks
' V0.6
' nicht fertig
' 03.04.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de
'@folder(Kennzeichen.SPS_RACK)

Option Explicit

Public Sub SPS_RackBMK()
    ' Erzeugen des gesamten Anlagen und Ortskennzeichen für SPS-Rack
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim TabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim SpalteStationsnummer As String
    
    Dim spalteEinbauortRack As String
    Dim spalteRackAnlagenkennzeichen As String
    Dim spalteAnlagenkennzeichen As String
    Dim sSpalteRackBMKperSignal As String
    Dim sSpalteStationPerSignal As String
    Dim dataAnlagenkennzeichen As String
    Dim dataRackAnlagenkennzeichen As String
    Dim iSpalteRackBMKperSignal As Long
    Dim iSpalteStationPerSignal As Long
    Dim tmpSpalteStationsnummer As Long
    Dim answer As Long
    Dim iSignal As Long
    Dim iSearchNumber As Long
    Dim EinbauorteData As New cEinbauorte
    Dim sResult As cEinbauorte
    Dim ExcelConfig As New cExcelConfig
    
    iSearchNumber = 0
    tmpSpalteStationsnummer = 0
      
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](TabelleDaten)
   
    Application.ScreenUpdating = False

    'read installation locations
    Set EinbauorteData = readEinbauorte(TabelleDaten)

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        spalteAnlagenkennzeichen = ExcelConfig.Anlage
        SpalteStationsnummer = ExcelConfig.Stationsnummer
        spalteEinbauortRack = ExcelConfig.SPSRackEinbauort
        spalteRackAnlagenkennzeichen = ExcelConfig.SPSRackAnlage
        sSpalteRackBMKperSignal = ExcelConfig.SPSRackBMKSignal_1
        sSpalteStationPerSignal = ExcelConfig.StationsnummerSignal_1
        iSpalteRackBMKperSignal = SpaltenBuchstaben2Int(sSpalteRackBMKperSignal)
        iSpalteStationPerSignal = SpaltenBuchstaben2Int(sSpalteStationPerSignal)
 
        answer = MsgBox("Spalte BU Stationsnummern und Einbauorte schon geprüft?", vbQuestion + vbYesNo + vbDefaultButton2, "Prüfung der Daten")
        'Prüfe Stationsnummer
        If answer = vbYes Then
 
            ' Spaltenbreiten anpassen
            ActiveSheet.Columns.Item(spalteRackAnlagenkennzeichen).Select
            Selection.ColumnWidth = 35
         
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' lesen von Feld Anlagenkennzeichen, führende Leerzeichen entfernen
                dataAnlagenkennzeichen = LTrim$(.Cells.Item(i, spalteAnlagenkennzeichen))
                ' Prüfe ob Stationsnummer mit Eintrag
                If .Cells.Item(i, SpalteStationsnummer) <> vbNullString Then
                    ' Anlagenkennzeichen ermitteln
                    dataRackAnlagenkennzeichen = "=" + Left$(dataAnlagenkennzeichen, InStr(1, dataAnlagenkennzeichen, "."))
                    If Len(.Cells.Item(i, SpalteStationsnummer)) = 1 Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen & "A.S0" & .Cells.Item(i, SpalteStationsnummer)
                    Else
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen & "A.S" & .Cells.Item(i, SpalteStationsnummer)
                    End If
                    ' wenn Einbauort nicht leer
                    If .Cells.Item(i, spalteEinbauortRack) <> vbNullString Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen & "+" & .Cells.Item(i, spalteEinbauortRack)
                    End If
                    ' Daten schreiben
                    .Cells.Item(i, spalteRackAnlagenkennzeichen) = dataRackAnlagenkennzeichen
                Else
                    ' Daten leeren
                    .Cells.Item(i, spalteRackAnlagenkennzeichen) = vbNullString
                End If
                For iSignal = 1 To 6
                    tmpSpalteStationsnummer = iSpalteStationPerSignal + (14 * (iSignal - 1))
                    iSearchNumber = .Cells.Item(i, tmpSpalteStationsnummer)
                    If iSearchNumber <> 0 Then   'nur weiter wenn Stationsnummer nicht leer
                        'Suchen nach den passenden Einbauort zur Station
                        Set sResult = Nothing
                        Set sResult = EinbauorteData.searchEinbauortDataset(iSearchNumber)
                    
                        If Not (sResult Is Nothing) Then ' prüfen ob etwas zurück kam
                            If sResult.Count > 0 Then ' nur weiter wenn Datensatz wirklich da
                                ' lesen von Feld Anlagenkennzeichen, führende Leerzeichen entfernen
                                dataAnlagenkennzeichen = LTrim$(.Cells.Item(i, spalteAnlagenkennzeichen))
                                ' Anlagenkennzeichen ermitteln
                                dataRackAnlagenkennzeichen = "=" + Left$(dataAnlagenkennzeichen, InStr(1, dataAnlagenkennzeichen, "."))
                                If Len(.Cells.Item(i, tmpSpalteStationsnummer)) = 1 Then
                                    dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen & "A.S0" & .Cells.Item(i, tmpSpalteStationsnummer)
                                Else
                                    dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen & "A.S" & .Cells.Item(i, tmpSpalteStationsnummer)
                                End If
                    
                                .Cells.Item(i, iSpalteRackBMKperSignal + (14 * (iSignal - 1))) = dataRackAnlagenkennzeichen & "+" & sResult.Item(1).Einbauort
                            End If
                        Else
                            ' makiere fehlende / falsche Steckplatz Daten
                            .Cells.Item(i, iSpalteRackBMKperSignal + (14 * (iSignal - 1))).Interior.ColorIndex = 3
                        End If
                    End If
                Next iSignal
            Next i
        Else
            MsgBox "Bitte Skript Stationsnummer ausführen und prüfen!"
        End If
    
    End With

End Sub

