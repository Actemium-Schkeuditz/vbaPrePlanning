Attribute VB_Name = "mConfigPLC"
' Skript zum Übertragen der SPS Konifguration nach Excel
' V0.2
' 17.02.2020
' update
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.SPS-Konfig)
'todo Config Datei muss noch mit Werten befüllt werden

Option Explicit

Public Sub ConfigPLC()
    'writes PLC Config to Excel sheets
    Dim i As Long
    Dim OffsetSlot As Integer

    ' Class einbinden
    Dim sData As New cBelegung
    Dim dataConfig As New cPLCconfig
    Dim dataConfigSort As cPLCconfig
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchConfig As New cKanalBelegungen
    
    ' Workbook
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tablename As String
    Dim zeilenanzahl As Long
    Dim spalteStationsnummer As String
    Dim spalteKartentyp As String
    
 
    ' Tabellen definieren
    tablename = "EplSheet"
    spalteStationsnummer = "BU"                  'erste Spalte der Anschlüsse
    spalteKartentyp = "BY"
    
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tablename)
   
    Application.ScreenUpdating = False

    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tablename, dataKanaele, spalteStationsnummer, spalteKartentyp
    'todo Belegungsdaten ermitteln pro Station für jeden Steckplatz den ersten Kanal mit Adresse und Typ ermitteln
   
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = dataKanaele.returnStation
    
    '##### Suche nach allen verwendeten Kartentypen
    Dim iKartentyp As Collection
   
    
    '####### bearbeiten der Daten #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    
    ' Variablen zum Schreiben
      Dim rTable As Range
        Set rTable = Range("A1")
        'Dim wdata As New cPLCconfigData
    
    For Each pStation In iStation
        ' suchen der Datensätze pro Station
        Set dataSearchStation = dataKanaele.searchDatasetPerStation(pStation)
        'Set dataSearchStation = dataKanaele.returnAllSlotsPerRack
        Set dataSearchConfig = dataSearchStation.returnAllSlotsPerRack
        'dataSearchStation.returnAllSlotsPerRack
    
        ' Übertragen der Daten
        Dim iAdressOutput As Long
        Dim iAdressInput As Long
        Set dataConfig = Nothing                 ' Rücksetzen der Datensammlung
        For Each sData In dataSearchConfig
            'sdata.Stationsnummer
            'sdata.Steckplatz
            'sdata.Kartentyp
            'sdata.Adress
            If sData.Kartentyp.InputAdressLength > 0 Or sData.Kartentyp.InputAdressDiagnosticLength > 0 Then
                iAdressInput = ExtractNumber(sData.Adress)
            End If
            If sData.Kartentyp.OutputAdressLength > 0 Or sData.Kartentyp.OutputAdressDiagnosticLength > 0 Then
                iAdressOutput = ExtractNumber(sData.Adress)
            End If
            dataConfig.Add sData.Stationsnummer, sData.Steckplatz, sData.Kartentyp.Kartentyp, sData.Key, iAdressInput, iAdressOutput
        Next
    
        ' Sortieren der Steckplätze
        Set dataConfigSort = dataConfig.Sort
     
        ' Tabelle für jede Station schreiben
      'dataConfigSort.writePLCConfigToExcel "Station_" & pStation
            'todo copy Data works not
           ' newExcelFile "SPS_CONFIG_3.xlsx", "config"
            
            'CopySheetToClosedWB "SPS_CONFIG_3.xlsx", "Station_" & pStation
    Next

    ' copy PLS config to file
          '  xCopyWorksheets "config"
End Sub



Sub readConfigFromSavedFile()
    'works fine
    Dim sfolder As String
    Dim sFileNameConfig As String
    Dim bConfigFileIsNew As Boolean
    Dim bConfigFileExist As Boolean
  
    ''''''''''''' Testen lesen in andere Exceltabelle
    sfolder = "config"
    sFileNameConfig = "SPSConfig.xlsm"


    ' Lege Datei an wenn nicht da
    bConfigFileIsNew = newExcelFile(sFileNameConfig, sfolder)

    ' wenn die Datei nicht angelegt wurde prüfe ob es diese schon gibt
    If bConfigFileIsNew = False Then
        'prüfe ob Datei vorhanden

        bConfigFileExist = fileExist(sFileNameConfig, sfolder)

        If bConfigFileExist Then
            'hier dann weiter wenn die Datei schon da ist
            'MsgBox "Datei gibt es schon"
            Dim Result As String
            Result = ReadSecondExcelFile(sFileNameConfig, sfolder)
   
            MsgBox Result
    
    
        ElseIf bConfigFileIsNew = True Then
            'todo hier weiter wenn Datei neu
        Else
            MsgBox "Fehler mit der SPSConfig.xlsx"
  
        End If

    End If

End Sub

Sub testenLesenConfig()
' Class einbinden
Dim dataConfig As New cPLCconfig
Dim tablename As String

    tablename = "Station_1"
    dataConfig.ReadPLCConfigData tablename

   MsgBox "daten gelesen"

tablename = "Station_16"
    dataConfig.ReadPLCConfigData tablename
   MsgBox "daten gelesen"
End Sub



