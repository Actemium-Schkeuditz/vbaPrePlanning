Attribute VB_Name = "SPSZuweisenKanal"
' Skript zur Ermittlung der SPS Kanäle
' V0.1
' nicht fertig
' 10.02.2020
'
'
' Christian Langrock
' christian.langrock@actemium.de



Option Explicit

Public Sub SPSZuweisenKanal()
 Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim y As Integer
    Dim it As Variant


 Dim spalteStationsnummer As String
 Dim spalteKartentyp As String
 
    Dim spalteIntStart As Integer
    
     ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearch As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    
     ' Tabellen definieren
    tabelleDaten = "EplSheet"
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets(tabelleDaten)
    spalteStationsnummer = "BU"                    'erste Spalte der Anschlüsse
    spalteKartentyp = "BY"
    ' Tabelle mit Daten bearbeiten
    With ws1
    
    
    ' Spaltenbreiten anpassen
        'ThisWorkbook.Worksheets(tabelleDaten).Activate
        ws1.Activate

        Application.ScreenUpdating = False
 
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt

        'dataKanaele.Add 10, "StrindDF", 4

        ' lesen der belegten Kanäle aus Excel Tabelle
        dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele, spalteStationsnummer, spalteKartentyp
   ' MsgBox "Daten gelesen"
    
    
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = returnStation(dataKanaele)
    
    '### Sortieren test ####
    Dim sortierung As cBelegung
    Dim dataSort As cKanalBelegungen

    Set dataSort = dataKanaele.Sort
    'dataKanaele.Sort
    
    For Each sortierung In dataSort
        Debug.Print sortierung.Stationsnummer; vbTab; sortierung.Kartentyp; vbTab; sortierung.Kanal
    Next
    
    
    '####### Suchen nach Kartentyp und Stationsnummer   ########
    ' hier nur Fest Station 13 und "IFM IO-LINK"
  
        dataSearch.searchKanalBelegungenKartentyp iStation, "IFM IO-LINK", dataSort
    'Next
    'MsgBox "Suche durchgeführt"
    '####### zuweisen der Kanäle #######
    ' hier fest IFM IO-Link
    dataResult.zuweisenKanal dataSearch
    
    'MsgBox "Zuweisung durchgeführt"
    
    '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
    
    Dim sData As New cBelegung
    For Each sData In dataResult
        For i = 3 To zeilenanzahl
            ' suchen nach dem pasenden KWSBMK
            If .Cells(i, "B") = sData.KWSBMK Then
                .Cells(i, "CA") = sData.Steckplatz
                .Cells(i, "CB") = sData.Kanal
                'MsgBox "Kanal geschrieben: " & sData.Kanal
                End If
            
        Next i
    Next
    
    End With
    

End Sub
