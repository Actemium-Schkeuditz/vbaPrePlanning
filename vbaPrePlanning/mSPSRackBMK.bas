Attribute VB_Name = "mSPSRackBMK"
' Skript zur Ermittlung der Anlagen und Ortskennzeichen der IO-Racks
' V0.7
' angepasst Anlagenkennzeichen ohne Punkt
' 28.04.2020
'
' Christian Langrock
' christian.langrock@actemium.de
'@folder(Kennzeichen.SPS_RACK)

Option Explicit

Public Sub SPS_RackBMK()
    ' Erzeugen des gesamten Anlagen und Ortskennzeichen f�r SPS-Rack
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
    Dim dataRackAnlagenkennzeichenFull As String
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

    Set ws1 = Worksheets.[_Default](TabelleDaten)
   
    Application.ScreenUpdating = False

    'read installation locations
    Set EinbauorteData = readEinbauorte(TabelleDaten)

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gez�hlt
        'MsgBox zeilenanzahl

        spalteAnlagenkennzeichen = ExcelConfig.Anlage
        SpalteStationsnummer = ExcelConfig.Stationsnummer
        spalteEinbauortRack = ExcelConfig.SPSRackEinbauort
        spalteRackAnlagenkennzeichen = ExcelConfig.SPSRackAnlage
        sSpalteRackBMKperSignal = ExcelConfig.SPSRackBMKSignal
        sSpalteStationPerSignal = ExcelConfig.StationsnummerSignal
        iSpalteRackBMKperSignal = SpaltenBuchstaben2Int(sSpalteRackBMKperSignal)
        iSpalteStationPerSignal = SpaltenBuchstaben2Int(sSpalteStationPerSignal)
 
        answer = MsgBox("Spalte BU Stationsnummern und Einbauorte schon gepr�ft?", vbQuestion + vbYesNo + vbDefaultButton2, "Pr�fung der Daten")
        'Pr�fe Stationsnummer
        If answer = vbYes Then
 
            ' Spaltenbreiten anpassen
            ActiveSheet.Columns.Item(spalteRackAnlagenkennzeichen).Select
            Selection.ColumnWidth = 35
         
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' lesen von Feld Anlagenkennzeichen, f�hrende Leerzeichen entfernen
                If InStr(2, .Cells.Item(i, spalteAnlagenkennzeichen), ".") Then
                    dataAnlagenkennzeichen = LTrim$(.Cells.Item(i, spalteAnlagenkennzeichen))
                    dataRackAnlagenkennzeichen = "=" + Left$(dataAnlagenkennzeichen, InStr(1, dataAnlagenkennzeichen, "."))
                Else
                    dataAnlagenkennzeichen = .Cells.Item(i, spalteAnlagenkennzeichen)
                    dataRackAnlagenkennzeichen = "=" & dataAnlagenkennzeichen & "."
                End If
                ' Pr�fe ob Stationsnummer mit Eintrag
                If .Cells.Item(i, SpalteStationsnummer) <> vbNullString Then
                    ' Anlagenkennzeichen ermitteln
                    If Len(.Cells.Item(i, SpalteStationsnummer)) = 1 Then
                        dataRackAnlagenkennzeichenFull = dataRackAnlagenkennzeichen & "A.S0" & .Cells.Item(i, SpalteStationsnummer)
                    Else
                        dataRackAnlagenkennzeichenFull = dataRackAnlagenkennzeichen & "A.S" & .Cells.Item(i, SpalteStationsnummer)
                    End If
                    ' wenn Einbauort nicht leer
                    If .Cells.Item(i, spalteEinbauortRack) <> vbNullString Then
                        dataRackAnlagenkennzeichenFull = dataRackAnlagenkennzeichenFull & "+" & .Cells.Item(i, spalteEinbauortRack)
                    End If
                    ' Daten schreiben
                    .Cells.Item(i, spalteRackAnlagenkennzeichen) = dataRackAnlagenkennzeichenFull
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
                    
                        If Not (sResult Is Nothing) Then ' pr�fen ob etwas zur�ck kam
                            If sResult.Count > 0 Then ' nur weiter wenn Datensatz wirklich da
                                If Len(.Cells.Item(i, tmpSpalteStationsnummer)) = 1 Then
                                    dataRackAnlagenkennzeichenFull = dataRackAnlagenkennzeichen & "A.S0" & .Cells.Item(i, tmpSpalteStationsnummer)
                                Else
                                    dataRackAnlagenkennzeichenFull = dataRackAnlagenkennzeichen & "A.S" & .Cells.Item(i, tmpSpalteStationsnummer)
                                End If
                                .Cells.Item(i, iSpalteRackBMKperSignal + (14 * (iSignal - 1))) = dataRackAnlagenkennzeichenFull & "+" & sResult.Item(1).Einbauort
                            End If
                        Else
                            ' makiere fehlende / falsche Steckplatz Daten
                            .Cells.Item(i, iSpalteRackBMKperSignal + (14 * (iSignal - 1))).Interior.ColorIndex = 3
                        End If
                    End If
                Next iSignal
            Next i
        Else
            MsgBox "Bitte Skript Stationsnummer ausf�hren und pr�fen!"
        End If
    End With
End Sub

