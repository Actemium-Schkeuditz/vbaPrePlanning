Attribute VB_Name = "mAssignAddresses"
' Funktion zur Ermittlung der SPS Adressen
' V0.3
' nicht fertig
' 18.02.2020
'diverse Fehler müssen abgefangen werden, Offset der Karten fehlt noch
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Adressen)

Option Explicit

Public Sub assignAddresses()
    'todo testen der Adressvergabe
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
    Dim dataPLCConfig As New cPLCconfig          'Config from File
    Dim dataConfigPerPLCTyp As New cPLCconfig
    
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
    Dim dataSort As New cKanalBelegungen         'Ergebnis der Sortierung
    'Set dataSort = dataKanaele.Sort
        
    
    
    For Each pStation In iStation
        '### suchen der Signale für die Station
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        '### lesen der PLC Konfiguartionsdaten ######
        Set dataPLCConfig = Nothing
        dataPLCConfig.ReadPLCConfigData "Station_" & pStation
        '### ermitteln der Adressen pro Station und schreiben der Datensätze
        Set dataResult = dataSearchPlcTyp.sumAdresses(pStation, dataPLCConfig)
       
        '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
        dataResult.writeDatsetsToExcel tabelleDaten
        'Next
    Next
    
    
End Sub

