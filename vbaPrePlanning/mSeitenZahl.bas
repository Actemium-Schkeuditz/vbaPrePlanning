Attribute VB_Name = "mSeitenZahl"
Option Explicit
' Skript zum schreiben der Seitenzahlen
' V0.5
' 12.03.2020
' neu Funktion zum schreiben der Seitenzahl
'
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.Seitenzahl)

Public Sub SeitenZahlschreiben()
    ' neu da alte Berechnung fehlerhaft

    Dim tabelleDaten As String
    Dim iSignal As Integer
    Dim iSeite As Long
    Dim AnlageOld As String
    Dim EinbauortOld As String
      
    ' Class einbinden
    Dim dataKanaele As New cKanalBelegungen
    Dim dataKanaeleSignal As New cKanalBelegungen
    Dim dataKanaelePneumatik As New cKanalBelegungen
    Dim dataKanaeleElektrik As New cKanalBelegungen
    Dim dataKanaelePneumatikSort As New cKanalBelegungen
    Dim dataKanaeleElektrikSort As New cKanalBelegungen
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
 
    On Error GoTo ErrorHandle
 
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    iSignal = 1

    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele
              
    '##### suchen nach den Datensätzen mit dem Signal 1
    Set dataKanaeleSignal = dataKanaele.searchDatasetSignal(iSignal)
              
    '#### suchen nach den Datensätzen die pneumatisch sind
    Set dataKanaelePneumatik = dataKanaeleSignal.searchDatasetPlcTyp("FESTO MPA")
    '#### suchen nach den Datensätzen die nicht pneumatisch sind
    Set dataKanaeleElektrik = dataKanaeleSignal.excludeDatasetPlcTyp("FESTO MPA")
    '#### sortieren nach Anlage und Einbauort
    Set dataKanaeleElektrikSort = dataKanaeleElektrik.SortForPageNumbers
    Set dataKanaelePneumatikSort = dataKanaelePneumatik.SortForPageNumbers
    
    '##### Seiten zuweisen elektrisch
    AnlageOld = vbNullString
    EinbauortOld = vbNullString
    iSeite = 1
    
    For Each sData In dataKanaeleElektrikSort
      
        If sData.Anlage = AnlageOld Then
            If sData.Einbauort = EinbauortOld Then
                iSeite = iSeite + 1
            Else
                iSeite = 1
            End If
        Else
            iSeite = 1
        End If
              
        AnlageOld = sData.Anlage
        EinbauortOld = sData.Einbauort
        sData.Seite = iSeite
        rData.AddDataSet sData
    Next
    
    '##### Seiten zuweisen pneumatisch
    AnlageOld = vbNullString
    EinbauortOld = vbNullString
    iSeite = 1
    
    For Each sData In dataKanaelePneumatikSort
      
        If sData.Anlage = AnlageOld Then
            If sData.Einbauort = EinbauortOld Then
                iSeite = iSeite + 1
            Else
                iSeite = 1
            End If
        Else
            iSeite = 1
        End If
        
        AnlageOld = sData.Anlage
        EinbauortOld = sData.Einbauort
        sData.Seite = iSeite
        rData.AddDataSet sData
    Next
    
    '#### Daten schreiben
    rData.writePageNumbersToExcel tabelleDaten

BeforeExit:
    ' Set rCell = Nothing
    'Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler im Modul Seitenzahl.", vbCritical, "Error"
    Resume BeforeExit

End Sub

