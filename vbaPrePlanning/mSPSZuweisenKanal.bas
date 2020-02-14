Attribute VB_Name = "mSPSZuweisenKanal"
' Skript zur Ermittlung der SPS Kanäle
' V0.2
' nicht fertig
' 13.02.2020
'diverse Fehler müssen abgefangen werden, Offset der Kartenn fehlt noch
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Kanalbelegung)

Option Explicit

Public Sub SPSZuweisenKanal()

    Dim tabelleDaten As String
    Dim i As Long
    Dim spalteStationsnummer As String
    Dim spalteKartentyp As String
    Dim offsetSlot As Integer
 
    
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    spalteStationsnummer = "BU"                  'erste Spalte der Anschlüsse
    spalteKartentyp = "BY"
    
    
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele, spalteStationsnummer, spalteKartentyp
    
    
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = dataKanaele.returnStation
    
    '##### Suche nach allen verwendeten Kartentypen
    Dim iKartentyp As Collection
    'Set iKartentyp = dataKanaele.returnKartentyp
    
    '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
    Dim sortierung As cBelegung
    Dim dataSort As New cKanalBelegungen             'Ergebnis der Sortierung
    'Set dataSort = dataKanaele.Sort
        
  
    '####### zuweisen der Kanäle #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    
    
    
    For Each pStation In iStation
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        offsetSlot = 0
        For Each pKartentyp In iKartentyp
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcTyp(pKartentyp)
            Set dataResult = dataSearchPlcTyp.zuweisenKanal(offsetSlot, pKartentyp)
            'MsgBox "Zuweisung durchgeführt"
            'TODO Offset verbessern
            'todo Behandlung Not-Aus und Festo CPX-8DE-D wegen doppel Stecker
            offsetSlot = dataResult.returnLastSlotNumber
            offsetSlot = offsetSlot + 1
            '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
            dataResult.writeDatsetsToExcel tabelleDaten
        Next
    Next
    
   
    
    
    '##### Anschlüsse zuordnen ####
      SPS_KartenAnschluss
    '##### SPS Karten BMK erzeugen ####
    SPS_BMK
    

End Sub


