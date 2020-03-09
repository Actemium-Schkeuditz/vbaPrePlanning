Attribute VB_Name = "mSPSZuweisenKanal"
' Skript zur Ermittlung der SPS Kanäle
' V0.4
' nicht fertig
' 02.03.2020
'diverse Fehler müssen abgefangen werden, Offset der Kartenn fehlt noch
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Kanalbelegung)

Option Explicit

Public Sub SPSZuweisenKanal()

    Dim tabelleDaten As String
    Dim i As Long
    'Dim spalteStationsnummer As String
   ' Dim spalteKartentyp As String
    Dim OffsetSlot As Integer
 
    Dim iInputAdress As Long
    Dim iOutputAdress As Long
    Dim PLCcardTyp As String
    
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    Dim dataResultAdress As New cKanalBelegungen
    Dim dataPLCConfig As New cPLCconfig          'Config from File
     Dim dataPLCConfigStation As New cPLCconfig          'Config for Work
    Dim dataConfigPerPLCTyp As New cPLCconfig
    Dim dataPLCConfigResult As New cPLCconfig
    Dim dataPLCConfigResultOutput As New cPLCconfig
  
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    'spalteStationsnummer = "BU"                  'erste Spalte der Anschlüsse
    'spalteKartentyp = "BY"
    
    iInputAdress = 0
    iOutputAdress = 0
    
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele ', spalteStationsnummer, spalteKartentyp
    
    
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = dataKanaele.returnStation
    
    '##### Suche nach allen verwendeten Kartentypen
    Dim iKartentyp As Collection
    
    '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
    'Dim sortierung As cBelegung
    Dim dataSort As New cKanalBelegungen         'Ergebnis der Sortierung
  
    '####### zuweisen der Kanäle #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    Dim iInputStartAdress As Long
    Dim iOutputStartAdress As Long
    Dim iMPAAnschlussplatte As Long
    Dim iSteckplatzMPA As Long
    Dim iKanalMPA As Long
      '##### Auslesen der Startadresse für die erste Station ######
    Set dataPLCConfig = readXMLFile
      
    iInputStartAdress = dataPLCConfig.returnFirstInputAdressePLCStation(1)
    iOutputStartAdress = dataPLCConfig.returnFirstOutputAdressePLCStation(1)
    
    For Each pStation In iStation
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        '### lesen der PLC Konfiguartionsdaten ######
        Set dataPLCConfigStation = Nothing
        Set dataPLCConfigResultOutput = Nothing
        Set dataResultAdress = Nothing
        Set dataPLCConfigStation = dataPLCConfig.returnDatasetPerStation(pStation)
        iSteckplatzMPA = 0
        iKanalMPA = 0
        PLCcardTyp = dataPLCConfigStation.Item(1).Kartentyp.PLCtyp
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        OffsetSlot = 0                           'starten mit Slot 0
        iMPAAnschlussplatte = 0
        For Each pKartentyp In iKartentyp
            Set dataConfigPerPLCTyp = Nothing
            Set dataConfigPerPLCTyp = dataPLCConfigStation.returnDatasetPerSlottyp(pStation, pKartentyp)
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcTyp(pKartentyp)
            Set dataResult = dataSearchPlcTyp.zuweisenKanal(OffsetSlot, pKartentyp, dataConfigPerPLCTyp)
            OffsetSlot = dataResult.returnLastSlotNumber
            ' Korrektur FESTO Ventilinsel
           Set dataResult = dataResult.correctFestoMPA(iMPAAnschlussplatte, iSteckplatzMPA, iKanalMPA)
            
            ' adressieren
            Set dataResultAdress = dataResult.AdressPerSlottyp(iInputStartAdress, iOutputStartAdress, pStation, pKartentyp)
            'Datensätze der Stationskonfiguration anhängen
            Set dataPLCConfigResult = dataPLCConfigResult.ConfigPLCToDataset(dataResultAdress)
            dataPLCConfigResultOutput.Append dataPLCConfigResult
           'Set dataPLCConfigResultOutput = dataPLCConfigResult.ConfigPLCToDataset(dataResultAdress)     'Datensätze der Stationskonfiguration anhängen
            
            OffsetSlot = OffsetSlot + 1
            '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
            dataResultAdress.writeDatsetsToExcel tabelleDaten
        Next
        'round up
        RoundUpPLCaddresses PLCcardTyp, iInputStartAdress, iOutputStartAdress
        
        dataPLCConfigResultOutput.writePLCConfigToExcel "Station_" & pStation
    Next
    
    

    
    
    '##### Anschlüsse zuordnen ####
    SPS_KartenAnschluss
    '##### SPS Karten BMK erzeugen ####
    SPS_BMK
    '##### CPX Daten Ergänzen ####
    CPXDatenErgaenzen
    
    MsgBox "Zuweisen fertig"
    
    


End Sub


