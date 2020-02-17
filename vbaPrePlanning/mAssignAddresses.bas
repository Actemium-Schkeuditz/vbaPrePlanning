Attribute VB_Name = "mAssignAddresses"
' Skript zur Ermittlung der SPS Adressen
' V0.1
' nicht fertig
' 17.02.2020
'diverse Fehler m�ssen abgefangen werden, Offset der Kartenn fehlt noch
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Adressen)

Option Explicit

Public Sub assignAddresses()

    Dim tabelleDaten As String
    Dim i As Long
    Dim spalteStationsnummer As String
    Dim spalteKartentyp As String
    Dim OffsetSlot As Integer
 
    
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    Dim dataPLCConfig As New cPLCconfig         'Config from File
    Dim dataConfigPerPLCTyp As New cPLCconfig
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    spalteStationsnummer = "BU"                  'erste Spalte der Anschl�sse
    spalteKartentyp = "BY"
    
    
    '##### lesen der belegten Kan�le aus Excel Tabelle #####
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
        
  
    '####### zuweisen der Kan�le #######
    ' Durchlauf f�r jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    
    
    
    For Each pStation In iStation
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        '### lesen der PLC Konfiguartionsdaten ######
        Set dataPLCConfig = Nothing
        dataPLCConfig.ReadPLCConfigData "Station_" & pStation
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        OffsetSlot = 0  'starten mit Slot 0
        'todo ab hier �berarbeiten
        For Each pKartentyp In iKartentyp
            Set dataConfigPerPLCTyp = Nothing
            Set dataConfigPerPLCTyp = dataPLCConfig.returnDatasetPerSlottyp(pStation, pKartentyp)
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcTyp(pKartentyp)
           ' Set dataResult = dataSearchPlcTyp.zuweisenKanal(OffsetSlot, pKartentyp, dataConfigPerPLCTyp)
            'MsgBox "Zuweisung durchgef�hrt"
            'TODO Offset verbessern
            'todo Behandlung Not-Aus und Festo CPX-8DE-D wegen doppel Stecker
            OffsetSlot = dataResult.returnLastSlotNumber
            OffsetSlot = OffsetSlot + 1
            '####### Zur�ckschreiben der Daten in urspr�ngliche Excelliste #######
            dataResult.writeDatsetsToExcel tabelleDaten
        Next
    Next
    
    
    End Sub