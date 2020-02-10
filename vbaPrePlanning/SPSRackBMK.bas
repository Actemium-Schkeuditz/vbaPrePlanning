Attribute VB_Name = "SPSRackBMK"
' Skript zur Ermittlung der Anlagen und Ortskennzeichen der IO-Racks
' V0.4
' nicht fertig
' 07.02.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de

Option Explicit

Public Sub SPS_RackBMK()
    ' Erzeugen des gesamten Anlagen und Ortskennzeichen für SPS-Rack
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim y As Integer
    Dim spalteStationsnummer As String
    
    Dim spalteEinbauortRack As String
    Dim spalteRackAnlagenkennzeichen As String
    Dim spalteSPSKanal As String
    Dim spalteAnlagenkennzeichen As String
    Dim dataAnlagenkennzeichen As String
    Dim dataRackAnlagenkennzeichen As String
    Dim answer As Integer
      
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

        spalteAnlagenkennzeichen = "B"
        spalteStationsnummer = "BU"
        spalteEinbauortRack = "BV"
        spalteRackAnlagenkennzeichen = "BW"
          
 
        answer = MsgBox("Spalte BU Stationsnummern und Einbauorte schon geprüft?", vbQuestion + vbYesNo + vbDefaultButton2, "Prüfung der Daten")
        'Prüfe Stationsnummer
        If answer = vbYes Then
 
            ' Spaltenbreiten anpassen
            ActiveSheet.Columns.Item(spalteRackAnlagenkennzeichen).Select
            Selection.ColumnWidth = 35
    
        
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' lesen von Feld Anlagenkennzeichen, führende Leerzeichen entfernen
                dataAnlagenkennzeichen = LTrim(Cells(i, spalteAnlagenkennzeichen))
                ' Prüfe ob Stationsnummer mit Eintrag
                If .Cells(i, spalteStationsnummer) <> vbNullString Then
                    ' Anlagenkennzeichen ermitteln
                    dataRackAnlagenkennzeichen = "=" + Left(dataAnlagenkennzeichen, InStr(1, dataAnlagenkennzeichen, "."))
                    If Len(.Cells(i, spalteStationsnummer)) = 1 Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "A.S0" + .Cells(i, spalteStationsnummer)
                    Else
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "A.S" + .Cells(i, spalteStationsnummer)
                    End If
                    ' wenn Einbauort nicht leer
                    If .Cells(i, spalteEinbauortRack) <> vbNullString Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "+" + .Cells(i, spalteEinbauortRack)
                    End If
                    ' Daten schreiben
                    .Cells(i, spalteRackAnlagenkennzeichen) = dataRackAnlagenkennzeichen
                Else
                    ' Daten leeren
                    .Cells(i, spalteRackAnlagenkennzeichen) = vbNullString
                End If
            Next i
        Else
            MsgBox "Bitte Skript Stationsnummer ausführen und prüfen!"
        End If
    
    End With

End Sub

Public Sub EinbauorteSchreiben()
    
    ' lesen der Einbauorte aus der Exceltabelle und schreiben der Felder "Einbauort" und "Einbauort des SPS-Rack´s"
    Dim EinbauorteData As New cEinbauorte        'Klasse anlegen für Datenaustausch

    Dim tablennameEinbauorte As String
    Dim spalteKWS_BMK As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    'Dim ws2 As Worksheet
    Dim tabelleDaten As String
    Dim dataKWSBMK As String
    Dim spalteStationsnummer As String
    Dim spalteEinbauortRack As String
    Dim spalteEinbauort As String
    Dim sResult As String
    Dim iSearchNumber As Integer

    'Tabellenamen ermitteln
    'ToDo
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    spalteKWS_BMK = "B"
    spalteStationsnummer = "BU"
    spalteEinbauortRack = "BV"
    spalteEinbauort = "BQ"
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets(tabelleDaten)
      
   
    Application.ScreenUpdating = False

    ' Tabelle mit Planungsdaten auslesen
    With ws1
        dataKWSBMK = LTrim(Cells(3, spalteKWS_BMK))
    
        If dataKWSBMK <> vbNullString Then
            If Left(dataKWSBMK, 3) = "BAP" Then
                tablennameEinbauorte = "Einbauorte_BAP"
            ElseIf Left(dataKWSBMK, 4) = "SG01" Then
                tablennameEinbauorte = "Einbauorte_H02.SG01"
            ElseIf Left(dataKWSBMK, 4) = "HDMA" Then
                tablennameEinbauorte = "Einbauorte_H03.HDMA"
            ElseIf Left(dataKWSBMK, 3) = "PPP" Then
                tablennameEinbauorte = "Einbauorte_MH04.PPP"
            ElseIf Left(dataKWSBMK, 5) = "SRN01" Then
                tablennameEinbauorte = "Einbauorte_MH04.SRN"
            ElseIf Left(dataKWSBMK, 5) = "TRP01" Or Left(dataKWSBMK, 5) = "TRP03" Then
                tablennameEinbauorte = "Einbauorte_MH03.KT1000"
            Else
                MsgBox "Keine passenden Daten mit Einbauorten gefunden, für KWS-BMK: " & dataKWSBMK
                tablennameEinbauorte = vbNullString
                Exit Sub                         ' hier dann Abbruch der ganzen Funktion
            End If
        Else
            MsgBox "Fehler in Daten, KWS-BMK erwartet"
        End If
    
    
        ' hier einlesen der Daten aus der Exceltabelle Einbauorte für die einzelnen Anlagen
        EinbauorteData.ReadExcelDataToCollection tablennameEinbauorte, EinbauorteData

        ' suchen nach Einbauort passend zur Stationsnummer
        'iSearchNumber = 10
        sResult = "leer"

        ' Spaltenbreiten anpassen
        ThisWorkbook.Worksheets(tabelleDaten).Activate
        ActiveSheet.Columns.Item(spalteEinbauort).Select
        '.Columns.Item(spalteEinbauort).Select
        Selection.ColumnWidth = 15
        ActiveSheet.Columns.Item(spalteEinbauortRack).Select
        Selection.ColumnWidth = 15

        'Herausfinden der Anzahl der Zeilen im Blatt der Vorplanungsdaten
        zeilenanzahl = .Cells(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        For i = 3 To zeilenanzahl
            iSearchNumber = .Cells(i, spalteStationsnummer)

            'Suchen nach den passenden Einbauort zur Station
            sResult = EinbauorteData.searchEinbauort(iSearchNumber, EinbauorteData)
            ' Einbauort des SPS-Racks schreiben
            If .Cells(i, spalteEinbauortRack) = sResult And (Not Trim(sResult) = Empty) Then
                ' Wenn gleich dann grün einfärben
                .Cells(i, spalteEinbauortRack).Interior.ColorIndex = 35
            Else                                 ' sonst gelb einfärben
                .Cells(i, spalteEinbauortRack).Interior.ColorIndex = 6
            End If
            .Cells(i, spalteEinbauortRack) = sResult
            If (Left(sResult, 2) <> "S1" And Left(sResult, 2) <> "S2" And Left(sResult, 2) <> "S3" And Left(sResult, 2) <> "Sx" And Left(sResult, 2) <> "SX") Or (Trim(sResult) = Empty) Then
                ' Einbauort schreiben
                If .Cells(i, spalteEinbauort) = sResult Then
                    ' Wenn gleich dann grün einfärben
                    .Cells(i, spalteEinbauort).Interior.ColorIndex = 35
                Else                             ' sonst gelb einfärben
                    .Cells(i, spalteEinbauort).Interior.ColorIndex = 6
                End If
                .Cells(i, spalteEinbauort) = sResult
            Else
                ' makiere fehlende / falsche Steckplatz Daten
                .Cells(i, spalteEinbauort).Interior.ColorIndex = 3
                .Cells(i, spalteEinbauortRack).Interior.ColorIndex = 3
            End If
        Next i
    End With
    MsgBox "Daten gelesen geschrieben. Spalte Einbauort kontollieren"

End Sub

