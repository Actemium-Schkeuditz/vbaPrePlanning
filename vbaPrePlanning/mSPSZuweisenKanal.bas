Attribute VB_Name = "mSPSZuweisenKanal"
' Skript zur Ermittlung der SPS Kanäle
' V0.6
' Änderung MPA Kartenzählung und Steckplatzbezeichnung
' 03.04.2020
'diverse Fehler müssen abgefangen werden,
'
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Kanalbelegung)

Option Explicit

Public Sub SPSZuweisenKanal()

    Dim TabelleDaten As String
    Dim OffsetSlot As Integer
    Dim PLCTyp As String
    Dim PLCTypOld As String

    '####### zuweisen der Kanäle #######
    Dim pStation As Variant
    Dim pKartentyp As Variant
    Dim iInputStartAdress As Long
    Dim iOutputStartAdress As Long

    '##### Suche nach allen verwendeten Kartentypen
    Dim iKartentyp As Collection
    
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen
    Dim dataResult As New cKanalBelegungen
    Dim dataResultAdress As New cKanalBelegungen
    Dim dataPLCConfig As New cPLCconfig          'Config from File
    Dim dataPLCConfigStation As New cPLCconfig   'Config for Work
    Dim dataConfigPerPLCTyp As New cPLCconfig
    Dim dataPLCConfigResult As New cPLCconfig
    Dim dataPLCConfigResultOutput As New cPLCconfig
    Dim dataMPAconfig As New cFestoMPA
    Dim ExcelConfig As New cExcelConfig
    
    '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
    Dim dataSort As New cKanalBelegungen         'Ergebnis der Sortierung
       
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten
    
    'Startwerte setzen
    dataMPAconfig.reset
    PLCTypOld = vbNullString
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection TabelleDaten, dataKanaele
        
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = dataKanaele.returnStation
 
    '##### Auslesen der Startadresse für die erste Station ######
    Set dataPLCConfig = readXMLFile
    iInputStartAdress = dataPLCConfig.returnFirstInputAdressePLCStation(1)
    iOutputStartAdress = dataPLCConfig.returnFirstOutputAdressePLCStation(1)
    
    For Each pStation In iStation
        '### reset der Festo config Daten
        dataMPAconfig.reset
        '### suchen nach den Datensätzen der Station
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        '### lesen der PLC Konfiguartionsdaten ######
        Set dataPLCConfigStation = Nothing
        Set dataPLCConfigResultOutput = Nothing
        Set dataResultAdress = Nothing
        Set dataPLCConfigStation = dataPLCConfig.returnDatasetPerStation(pStation)
        '### PLC Typ ermitteln
        PLCTyp = dataSearchStation.Item(1).Kartentyp.PLCTyp
        ' Erkennen von Stationswechseln und dann aufrunden der Adressen
        If PLCTyp <> PLCTypOld And Not PLCTypOld = vbNullString Then
            RoundUpPLCaddresses iInputStartAdress, iOutputStartAdress
        End If
        PLCTypOld = PLCTyp
    
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        OffsetSlot = 0                           'starten mit Slot 0
 
        For Each pKartentyp In iKartentyp
            Set dataConfigPerPLCTyp = Nothing
            Set dataConfigPerPLCTyp = dataPLCConfigStation.returnDatasetPerSlottyp(pStation, pKartentyp)
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcModules(pKartentyp)
            Set dataResult = dataSearchPlcTyp.zuweisenKanal(OffsetSlot, pKartentyp, dataConfigPerPLCTyp)
            PLCTyp = dataResult.Item(1).Kartentyp.PLCTyp
            OffsetSlot = dataResult.returnLastSlotNumber
            ' Korrektur FESTO Ventilinsel
            Set dataResult = dataResult.correctFestoMPA(dataMPAconfig)
            ' adressieren
            Set dataResultAdress = dataResult.AdressPerSlottyp(iInputStartAdress, iOutputStartAdress, pStation, pKartentyp)
            'Datensätze der Stationskonfiguration anhängen
            Set dataPLCConfigResult = dataPLCConfigResult.ConfigPLCToDataset(dataResultAdress)
            dataPLCConfigResultOutput.Append dataPLCConfigResult
            
            '##### symbolische Adresse ermitteln #####
            dataResultAdress.symbolischeAdressen
            
            OffsetSlot = OffsetSlot + 1
            '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
            dataResultAdress.writeDatsetsToExcel TabelleDaten
        Next
        '### schreiben der Config Daten in eigenens Excel Sheet
        dataPLCConfigResultOutput.writePLCConfigToExcel "Station_" & pStation, PLCTyp
    Next
    
    '##### Anschlüsse zuordnen #####
    SPS_KartenAnschluss
    '##### SPS Karten BMK erzeugen #####
    SPS_BMK
    '##### CPX Daten Ergänzen #####
    CPXDatenErgaenzen
    '##### Seitenzahl schreiben #####
    SeitenZahlschreiben

    MsgBox "Zuweisen fertig"
    
End Sub
