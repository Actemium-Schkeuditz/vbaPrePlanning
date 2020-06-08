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
    Dim PLCtyp As String
    Dim PLCTypOld As String
    Dim iStationOld As Long

    '####### zuweisen der Kanäle #######
    Dim pStation As Variant
    Dim pKartentyp As Variant
    Dim iInputStartAdress As Long
    Dim iOutputStartAdress As Long
    Dim iInputAdressUsed As Long
    Dim iOutputAdressUsed As Long
    Dim iTmpInputAdressUsed As Long
    Dim iTmpOutputAdressUsed As Long

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
    Dim dataPLCOverview As New cPLCconfig
    Dim dataMPAconfig As New cFestoMPA
    Dim ExcelConfig As New cExcelConfig
    
    '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
    Dim dataSort As New cKanalBelegungen         'Ergebnis der Sortierung
       
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten
    
    'Startwerte setzen
    dataMPAconfig.reset
    PLCTypOld = vbNullString
    iStationOld = 0
    iInputAdressUsed = 0
    iOutputAdressUsed = 0
    Set dataPLCOverview = Nothing
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
        PLCtyp = dataSearchStation.Item(1).Kartentyp.PLCtyp
        
        ' Erkennen von Stationswechseln und dann aufrunden der Adressen
        If PLCtyp <> PLCTypOld And Not PLCTypOld = vbNullString Or (PLCtyp = "ET200SP" And iStationOld > 0 And pStation <> iStationOld) Then
            RoundUpPLCaddresses iInputStartAdress, iOutputStartAdress
        End If
        PLCTypOld = PLCtyp
    
        '### Sortieren nach Stationsnummer, Sortierkennung der Karte und KWS-BMK ####
        Set dataSort = dataSearchStation.Sort
        '##### Suche nach allen verwendeten Kartentypen
        Set iKartentyp = dataSort.returnKartentyp
        OffsetSlot = 0                           'starten mit Slot 0
 
        For Each pKartentyp In iKartentyp
            Set dataConfigPerPLCTyp = Nothing
            Set dataConfigPerPLCTyp = dataPLCConfigStation.returnDatasetPerSlottyp(pStation, pKartentyp)
            Set dataSearchPlcTyp = dataSort.searchDatasetPlcModules(pKartentyp)
            ' Prüfen ob Projekt MH04.TRP
            If Left(dataSearchPlcTyp.Item(1).KWSBMK, 4) = "TRP." And pStation = 3 Then
                Set dataResult = dataSearchPlcTyp
            Else
                Set dataResult = dataSearchPlcTyp.zuweisenKanal(OffsetSlot, pKartentyp, dataConfigPerPLCTyp)
            End If
            
            PLCtyp = dataResult.Item(1).Kartentyp.PLCtyp
            OffsetSlot = dataResult.returnLastSlotNumber
            ' Korrektur FESTO Ventilinsel
             ' Prüfen ob Projekt MH04.TRP
            If Left(dataSearchPlcTyp.Item(1).KWSBMK, 4) = "TRP." And pStation = 3 Then
                Set dataResult = dataResult
            Else
                Set dataResult = dataResult.correctFestoMPA(dataMPAconfig)
            End If
            ' adressieren
            iTmpInputAdressUsed = iInputStartAdress
            iTmpOutputAdressUsed = iOutputStartAdress
            Set dataResultAdress = dataResult.AdressPerSlottyp(iInputStartAdress, iOutputStartAdress, pStation, pKartentyp)
            ' Benutzte Adressen ermitteln
            iInputAdressUsed = iInputAdressUsed + iInputStartAdress - iTmpInputAdressUsed
            iOutputAdressUsed = iOutputAdressUsed + iOutputStartAdress - iTmpOutputAdressUsed
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
        dataPLCConfigResultOutput.writePLCConfigToExcel "Station_" & pStation, PLCtyp
        '# Stationsnummer sichern
        iStationOld = pStation
    Next
    
    '##### Anschlüsse zuordnen #####
    SPS_KartenAnschluss
    '##### SPS Karten BMK erzeugen #####
    SPS_BMK
    '##### CPX Daten Ergänzen #####
    CPXDatenErgaenzen
    '##### Seitenzahl schreiben #####
    SeitenZahlschreiben
    
    '##### Ausgabe Übersicht #####
    dataPLCOverview.Add 0, 0, vbNullString, vbNullString, iInputAdressUsed, iOutputAdressUsed
    dataPLCOverview.writePLCOverviewToExcel "Übersicht"

    MsgBox "Zuweisen fertig"
    
End Sub
