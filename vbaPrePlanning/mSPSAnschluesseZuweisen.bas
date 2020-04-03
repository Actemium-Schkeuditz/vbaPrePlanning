Attribute VB_Name = "mSPSAnschluesseZuweisen"
' Skript zur Ermittlung der SPS Anschlüsse
' V0.13
' 03.04.2020
' Such- und Schreibfunktion fehlerbehoben
' Einbindung ExcelConfig

' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Anschlüsse)

Option Explicit

Public Sub SPS_KartenAnschluss()

    Dim TabelleDaten As String
      
    ' Class einbinden
    Dim dataAnschluesse As New CSPSAnschluesse
    Dim dataKanaele As New cKanalBelegungen
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
    Dim dataResult As New cAnschluss
    Dim ExcelConfig As New cExcelConfig
    
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten
   
    '##### lesen der Anschlussdaten aus Excel Tabelle  #####
    dataAnschluesse.ReadExcelDataConnectionToCollection dataAnschluesse
        
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection TabelleDaten, dataKanaele
              
    For Each sData In dataKanaele
        
        Set dataResult = Nothing
        Set dataResult = dataAnschluesse.searchAnschluss(sData.Kartentyp.Kartentyp, sData.Kanal, dataAnschluesse)
        If Not sData.Kartentyp.PLCTyp = "FESTO MPA" Then
            sData.Anschluss1 = dataResult.Anschluss1
            sData.Anschluss2 = dataResult.Anschluss2
            sData.Anschluss3 = dataResult.Anschluss3
            sData.Anschluss4 = dataResult.Anschluss4
            sData.AnschlussM = dataResult.AnschlussM
            sData.AnschlussVS = dataResult.AnschlussVS
            
            rData.AddDataSet sData
        End If
    Next
    '#### Daten schreiben
    rData.writeDatsetsToExcel TabelleDaten
        
End Sub

