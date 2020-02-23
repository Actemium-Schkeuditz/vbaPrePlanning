Attribute VB_Name = "mSPSZuweisenKanal"
' Skript zur Ermittlung der SPS Kanäle
' V0.3
' nicht fertig
' 23.02.2020
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
    Dim OffsetSlot As Integer
 
    Dim iInputAdress As Long
    Dim iOutputAdress As Long
    
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    Dim dataResultAdress As New cKanalBelegungen
    Dim dataPLCConfig As New cPLCconfig          'Config from File
    Dim dataConfigPerPLCTyp As New cPLCconfig
    Dim dataPLCConfigResult As New cPLCconfig
  
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    spalteStationsnummer = "BU"                  'erste Spalte der Anschlüsse
    spalteKartentyp = "BY"
    
    iInputAdress = 0
    iOutputAdress = 0
    
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
    Dim dataSort As New cKanalBelegungen         'Ergebnis der Sortierung
    'Set dataSort = dataKanaele.Sort
        
  
    '####### zuweisen der Kanäle #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    Dim iInputStartAdress As Long
    Dim iOutputStartAdress As Long
    
    For Each pStation In iStation
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        '### lesen der PLC Konfiguartionsdaten ######
        Set dataPLCConfig = Nothing
        dataPLCConfig.ReadPLCConfigData "Station_" & pStation
        ' Auslesen der Startadresse für eine Station
        iInputStartAdress = dataPLCConfig.returnFirstInputAdressePLCStation(pStation)
        iOutputStartAdress = dataPLCConfig.returnFirstOutputAdressePLCStation(pStation)
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        OffsetSlot = 0                           'starten mit Slot 0
        For Each pKartentyp In iKartentyp
            Set dataConfigPerPLCTyp = Nothing
            Set dataConfigPerPLCTyp = dataPLCConfig.returnDatasetPerSlottyp(pStation, pKartentyp)
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcTyp(pKartentyp)
            Set dataResult = dataSearchPlcTyp.zuweisenKanal(OffsetSlot, pKartentyp, dataConfigPerPLCTyp)
            'MsgBox "Zuweisung durchgeführt"
            'TODO Offset verbessern
            'todo Behandlung Festo CPX-8DE-D wegen Doppelstecker
            OffsetSlot = dataResult.returnLastSlotNumber
            
            ' adressieren
            Set dataResultAdress = dataResult.AdressPerSlottyp(iInputAdress, iOutputAdress, pStation, pKartentyp)
            dataPLCConfigResult.ConfigPLCToDataset dataResultAdress 'Datensätze der Stationskonfiguration anhängen
            'Set dataPLCConfigResult = ConfigPLCToDataset(dataResult)
            ' ermitteln der Startadressen der einzelnen Steckplätze
            '   dataPLCConfigResult.sumAdressesPerSlot pStation, dataResult
            ' den Kanälen Adressen zuweisen
            'dataPLCConfigResult.sumAdresses
            ' (ConfigPLCToDataset(dataResult))
            dataPLCConfigResult.ConfigPLCToDataset dataResultAdress    'Datensätze der Stationskonfiguration anhängen
            
            OffsetSlot = OffsetSlot + 1
            '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
            dataResultAdress.writeDatsetsToExcel tabelleDaten
        Next
    Next
    
   
    
    
    '##### Anschlüsse zuordnen ####
    SPS_KartenAnschluss
    '##### SPS Karten BMK erzeugen ####
    SPS_BMK
    
    MsgBox "Zuweisen fertig"

End Sub



