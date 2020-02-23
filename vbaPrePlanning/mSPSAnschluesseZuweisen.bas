Attribute VB_Name = "mSPSAnschluesseZuweisen"
' Skript zur Ermittlung der SPS Anschlüsse
' V0.10
' teilweise getestet fertig
' 11.02.2020
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
      
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
    spalteSignal_1_Typ = "BX"                    'erste Spalte der Anschlüsse
   
    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Konvertierung Spaltenbuchstaben in INTEGER
        spalteIntStart = SpaltenBuchstaben2Int(spalteSignal_1_Typ)
   
        ' Spaltenbreiten anpassen
        'ThisWorkbook.Worksheets(tabelleDaten).Activate
        ws1.Activate

        Application.ScreenUpdating = False
 
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt

        ' lesen der Anschlussdaten aus Excel Tabelle
        dataAnschluesse.ReadExcelDataConnectionToCollection dataAnschluesse
    
        ' ******* ab hier suchen und schreiben der Daten
        ' suchen nach Anschlüssen passend zum Kartentyp und zum Kanal
        'iSearchKanal = 10
        'iSearchKartentyp = "ET200SP 4FDO"
        spalteOffset = 0
    
        ' Alle fünf Kanäle abarbeiten
        For y = 0 To 4
            spalteOffset = y * 12
            For i = 3 To zeilenanzahl
                iSearchKartentyp = .Cells.Item(i, spalteIntStart + 1).Value
                If iSearchKartentyp <> vbNullString And (.Cells.Item(i, spalteIntStart + spalteOffset + 4) <> vbNullString) Then ' wenn Kartentyp nicht leer dann auslesen und schreiben
                    iSearchKanal = .Cells.Item(i, spalteIntStart + spalteOffset + 4).Value
                    'Suchen nach dem passenden Datensatz passend zu Kartentyp und Kanal
                    dataSearch.searchAnschluss iSearchKartentyp, iSearchKanal, dataAnschluesse
               
                    ' schreiben der Daten
                    .Cells.Item(i, spalteIntStart + spalteOffset + 6) = dataSearch.Item(1).Anschluss1
                    .Cells.Item(i, spalteIntStart + spalteOffset + 7) = dataSearch.Item(1).Anschluss2
                    .Cells.Item(i, spalteIntStart + spalteOffset + 8) = dataSearch.Item(1).Anschluss3
                    .Cells.Item(i, spalteIntStart + spalteOffset + 9) = dataSearch.Item(1).Anschluss4
                    .Cells.Item(i, spalteIntStart + spalteOffset + 10) = dataSearch.Item(1).AnschlussM
                    .Cells.Item(i, spalteIntStart + spalteOffset + 11) = dataSearch.Item(1).AnschlussVS
                    
                    'Debug.Print dataSearch.Item(1).Kartentyp; dataSearch.Item(1).Kanal; vbTab; dataSearch.Item(1).Anschluss_1; vbTab; dataSearch.Item(1).Anschluss_2
                    dataSearch.Remove (1)
                End If
            Next i
        Next y
      
        ws1.Activate
    End With
End Sub

