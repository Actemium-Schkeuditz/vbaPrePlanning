Attribute VB_Name = "mEinbauorte"
' Skript zur Ermittlung der Anlagen und Ortskennzeichen der IO-Racks
' V0.4
' abgetrennt aus SPSRackBMK
' 11.02.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de
 '@folder(Kennzeichen.Einbauorte)
Option Explicit

Public Sub EinbauorteSchreiben()
   
    ' lesen der Einbauorte aus der Exceltabelle und schreiben der Felder "Einbauort" und "Einbauort des SPS-Rack´s"
    Dim EinbauorteData As New cEinbauorte        'Klasse anlegen für Datenaustausch

    Dim tablennameEinbauorte As String
    Dim spalteKWS_BMK As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    'Dim ws2 As Worksheet
    Dim tabelleDaten As String
    Dim dataKWSBMK As String
    Dim spalteStationsnummer As String
    Dim spalteEinbauortRack As String
    Dim spalteEinbauort As String
    Dim sResult As String
    Dim iSearchNumber As Long

    'Tabellenamen ermitteln
    'ToDo Einbauorte übertragen AL1400 und AL1402
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    spalteKWS_BMK = "B"
    spalteStationsnummer = "BU"
    spalteEinbauortRack = "BV"
    spalteEinbauort = "BQ"
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
      
   
    Application.ScreenUpdating = False

    ' Tabelle mit Planungsdaten auslesen
    With ws1
        dataKWSBMK = LTrim$(.Cells.Item(3, spalteKWS_BMK))
    
        If dataKWSBMK <> vbNullString Then
            If Left$(dataKWSBMK, 3) = "BAP" Then
                tablennameEinbauorte = "Einbauorte_BAP"
            ElseIf Left$(dataKWSBMK, 4) = "SG01" Then
                tablennameEinbauorte = "Einbauorte_H02.SG01"
            ElseIf Left$(dataKWSBMK, 4) = "HDMA" Then
                tablennameEinbauorte = "Einbauorte_H03.HDMA"
            ElseIf Left$(dataKWSBMK, 3) = "PPP" Then
                tablennameEinbauorte = "Einbauorte_MH04.PPP"
            ElseIf Left$(dataKWSBMK, 5) = "SRN01" Then
                tablennameEinbauorte = "Einbauorte_MH04.SRN"
            ElseIf Left$(dataKWSBMK, 5) = "TRP01" Then
                tablennameEinbauorte = "Einbauorte_MH03.TRP01"
                ElseIf Left$(dataKWSBMK, 5) = "TRP03" Then
                tablennameEinbauorte = "Einbauorte_MH03.TRP03"
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
        ThisWorkbook.Worksheets.[_Default](tabelleDaten).Activate
        ActiveSheet.Columns.Item(spalteEinbauort).Select
        '.Columns.Item(spalteEinbauort).Select
        Selection.ColumnWidth = 15
        ActiveSheet.Columns.Item(spalteEinbauortRack).Select
        Selection.ColumnWidth = 15

        'Herausfinden der Anzahl der Zeilen im Blatt der Vorplanungsdaten
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        For i = 3 To zeilenanzahl
            iSearchNumber = .Cells.Item(i, spalteStationsnummer)

            'Suchen nach den passenden Einbauort zur Station
            sResult = EinbauorteData.searchEinbauort(iSearchNumber, EinbauorteData)
            ' Einbauort des SPS-Racks schreiben
            If .Cells.Item(i, spalteEinbauortRack) = sResult And (Not Trim$(sResult) = Empty) Then
                ' Wenn gleich dann grün einfärben
                .Cells.Item(i, spalteEinbauortRack).Interior.ColorIndex = 35
            Else                                 ' sonst gelb einfärben
                .Cells.Item(i, spalteEinbauortRack).Interior.ColorIndex = 6
            End If
            .Cells.Item(i, spalteEinbauortRack) = sResult
            If (Left$(sResult, 2) <> "S1" And Left$(sResult, 2) <> "S2" And Left$(sResult, 2) <> "S3" And Left$(sResult, 2) <> "Sx" And Left$(sResult, 2) <> "SX") Or (Trim$(sResult) = Empty) Then
                ' Einbauort schreiben
                If .Cells.Item(i, spalteEinbauort) = sResult Then
                    ' Wenn gleich dann grün einfärben
                    .Cells.Item(i, spalteEinbauort).Interior.ColorIndex = 35
                Else                             ' sonst gelb einfärben
                    .Cells.Item(i, spalteEinbauort).Interior.ColorIndex = 6
                End If
                .Cells.Item(i, spalteEinbauort) = sResult
            Else
                ' makiere fehlende / falsche Steckplatz Daten
                .Cells.Item(i, spalteEinbauort).Interior.ColorIndex = 3
                .Cells.Item(i, spalteEinbauortRack).Interior.ColorIndex = 3
            End If
        Next i
    End With
    MsgBox "Daten gelesen und geschrieben. Spalte Einbauort kontollieren"

End Sub
