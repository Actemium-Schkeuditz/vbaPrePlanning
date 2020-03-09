Attribute VB_Name = "mSPSAnschluesseZuweisen"
' Skript zur Ermittlung der SPS Anschlüsse
' V0.11
' Verschiebung wegen neuer Stationsnummer
' 02.03.2020
' angepasst für Rubberduck
' Variablen umbenannt

' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Anschlüsse)

Option Explicit

Public Sub SPS_KartenAnschluss()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim y As Long
    
    'Dim sResult As New CSPSAnschluesse
    Dim iSearchKanal As Long
    Dim iSearchKartentyp As String
    
    Dim spalteSignal_1_Typ As String
    Dim spalteIntStart As Long
    Dim spalteOffset As Long

      
    ' Class einbinden
    Dim dataAnschluesse As New CSPSAnschluesse
    Dim dataSearch As New CSPSAnschluesse
    Dim dataResult As New CSPSAnschluesse
      
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
    spalteSignal_1_Typ = "CA"                    'erste Spalte der Anschlüsse "ACT.PLS.SIGNAL_1.KARTENTYP de_DE"
   
    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Konvertierung Spaltenbuchstaben in INTEGER
        spalteIntStart = SpaltenBuchstaben2Int(spalteSignal_1_Typ)
   
        ' Spaltenbreiten anpassen
        ws1.Activate

        Application.ScreenUpdating = False
 
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt

        ' lesen der Anschlussdaten aus Excel Tabelle
        dataAnschluesse.ReadExcelDataConnectionToCollection dataAnschluesse
    
        ' ******* ab hier suchen und schreiben der Daten
        ' suchen nach Anschlüssen passend zum Kartentyp und zum Kanal
        spalteOffset = 0
    
        ' Alle sechs Kanäle abarbeiten
        For y = 0 To 5
            spalteOffset = y * 14
            For i = 3 To zeilenanzahl
                iSearchKartentyp = .Cells.Item(i, spalteIntStart).Value
                If iSearchKartentyp <> vbNullString And (.Cells.Item(i, spalteIntStart + spalteOffset + 3) <> vbNullString) Then ' wenn Kartentyp nicht leer dann auslesen und schreiben
                    iSearchKanal = .Cells.Item(i, spalteIntStart + spalteOffset + 3).Value
                    'Suchen nach dem passenden Datensatz passend zu Kartentyp und Kanal
                    Set dataResult = Nothing
                    Set dataResult = dataSearch.searchAnschluss(iSearchKartentyp, iSearchKanal, dataAnschluesse)
                    If Not dataResult.Item(1).Kartentyp.PLCtyp = "FESTO MPA" Then
               
                        ' schreiben der Daten
                        .Cells.Item(i, spalteIntStart + spalteOffset + 5) = dataResult.Item(1).Anschluss1
                        .Cells.Item(i, spalteIntStart + spalteOffset + 6) = dataResult.Item(1).Anschluss2
                        .Cells.Item(i, spalteIntStart + spalteOffset + 7) = dataResult.Item(1).Anschluss3
                        .Cells.Item(i, spalteIntStart + spalteOffset + 8) = dataResult.Item(1).Anschluss4
                        .Cells.Item(i, spalteIntStart + spalteOffset + 9) = dataResult.Item(1).AnschlussM
                        .Cells.Item(i, spalteIntStart + spalteOffset + 10) = dataResult.Item(1).AnschlussVS
                        Else
                        
                    End If
                    'Debug.Print dataSearch.Item(1).Kartentyp; dataSearch.Item(1).Kanal; vbTab; dataSearch.Item(1).Anschluss_1; vbTab; dataSearch.Item(1).Anschluss_2
                    'dataSearch.Remove (1)
                End If
            Next i
        Next y
      
        ws1.Activate
    End With
End Sub



