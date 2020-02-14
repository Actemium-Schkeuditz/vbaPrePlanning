Attribute VB_Name = "mConfigPLC"
' Skript zum Übertragen der SPS Konifguration nach Excel
' V0.1
' 13.02.2020
' neu
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.SPS-Konfig)
'todo Config Datei muss noch mit Werten befüllt werden

Option Explicit

Public Sub ConfigPLC()

    Dim sFileNameConfig As String
    Dim sFolder As String
    Dim bConfigFileIsNew As Boolean
    Dim bConfigFileExist As Boolean
    Dim i As Long
    Dim offsetSlot As Integer

    ' Class einbinden
    Dim sdata As New cBelegung
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
   
    
    '####### zuweisen der Kanäle #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    
    ' Variablen zum Schreiben
      Dim rTable As Range
        Set rTable = Range("A1")
        Dim wData As New cPLCconfigData
    
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
        For Each sdata In dataSearchConfig
            'sdata.Stationsnummer
            'sdata.Steckplatz
            'sdata.Kartentyp
            'sdata.Adress
            If sdata.Kartentyp.InputAdressLength > 0 Or sdata.Kartentyp.InputAdressDiagnosticLength > 0 Then
                iAdressInput = ExtractNumber(sdata.Adress)
            End If
            If sdata.Kartentyp.OutputAdressLength > 0 Or sdata.Kartentyp.OutputAdressDiagnosticLength > 0 Then
                iAdressOutput = ExtractNumber(sdata.Adress)
            End If
            dataConfig.Add sdata.Stationsnummer, sdata.Steckplatz, sdata.Kartentyp.Kartentyp, iAdressInput, iAdressOutput
        Next
    
    
        ' Sortieren der Steckplätze
        Set dataConfigSort = dataConfig.Sort
     
        ' Tabelle mit Daten bearbeiten
        'With ws1
     With ThisWorkbook
            ' alte Daten löschen
            
            If WorksheetExist(("Station" & pStation), ws1) = True Then
            Application.DisplayAlerts = False
            .Sheets("Station" & pStation).Delete
            Application.DisplayAlerts = True
            End If
            'Worksheets anlegen
            .Sheets.Add after:=Sheets(Worksheets.Count)
            .ActiveSheet.Name = "Station" & pStation
            
            ' Daten einschreiben
      
        ThisWorkbook.Worksheets("Station" & pStation).Activate
        
      
        
     ' Tabellen kopf
     .ActiveSheet.Cells(1, 1) = "Stationsnummer"
     .ActiveSheet.Cells(1, 2) = "Steckplatz"
     .ActiveSheet.Cells(1, 3) = "Kartentyp"
     .ActiveSheet.Cells(1, 4) = "Eingangsadresse"
     .ActiveSheet.Cells(1, 5) = "Ausgangsadresse"
     .ActiveSheet.Cells(1, 6) = "Reservekanäle"
     .ActiveSheet.Cells(1, 7) = "Reservekanäle"
     .ActiveSheet.Cells(1, 8) = "ReserveSteckplätze"
     
            i = 2
            For Each wData In dataConfigSort
            .ActiveSheet.Cells(i, 1) = wData.Stationsnummer
            .ActiveSheet.Cells(i, 2) = wData.Steckplatz
            .ActiveSheet.Cells(i, 3) = wData.Kartentyp.Kartentyp
            .ActiveSheet.Cells(i, 4) = wData.FirstInputAdress
            .ActiveSheet.Cells(i, 5) = wData.FirstOutputAdress
            i = i + 1
            Next
            i = 0
   
      End With
    Next
    
    
    
    ''''''''''''' Testen lesen in andere Exceltabelle

    sFolder = "config"
    sFileNameConfig = "SPSConfig.xlsx"


    ' Lege Datei an wenn nicht da
    bConfigFileIsNew = newExcelFile(sFileNameConfig, sFolder)

    ' wenn die Datei nicht angelegt wurde prüfe ob es diese schon gibt
    If bConfigFileIsNew = False Then
        'prüfe ob Datei vorhanden

        bConfigFileExist = fileExist(sFileNameConfig, sFolder)

        If bConfigFileExist Then
            'todo hier dann weiter wenn die Datei schon da ist
            'MsgBox "Datei gibt es schon"
            Dim Result As String
            Result = ReadSecondExcelFile(sFileNameConfig, sFolder)
   
            MsgBox Result
    
    
        ElseIf bConfigFileIsNew = True Then
            'todo hier weiter wenn Datei neu
        Else
            MsgBox "Fehler mit der SPSConfig.xlsx"
  
        End If

    End If

End Sub


