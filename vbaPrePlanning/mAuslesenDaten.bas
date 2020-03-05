Attribute VB_Name = "mAuslesenDaten"
 Option Explicit

Public Sub AuslesenDaten()

    Dim tabelleDaten As String
    Dim dataKanaele As New cKanalBelegungen
    Dim sKanaele As New cBelegung
    Dim sortKanaele As New cKanalBelegungen
    Dim rData As New cKanalBelegungen
    
    Dim myblatt As Worksheet
    Dim Zaehler1 As Integer
    Dim spalteSig_1_Steckplatz As String
    Dim spalteSig_2_Steckplatz As String
    Dim spalteSig_1_Kanal As String
    Dim spalteSig_2_Kanal As String
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    
    ' Spalten definieren
    spalteSig_1_Steckplatz = "CC"
    spalteSig_2_Steckplatz = "CQ"
    spalteSig_1_Kanal = "CD"
    spalteSig_2_Kanal = "CD"
    
    

Set myblatt = Worksheets.[_Default](tabelleDaten)

With myblatt


  'For Zaehler1 = 3 To .Cells(Rows.Count, "A").End(xlUp).Row  'Gültigkeit in Spalte "A" von Zeile 3 bis Ende der Spalte

  
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele 'Auslesen der DAten aus Excelliste
    
    Set sortKanaele = dataKanaele.Sort 'Sortieren nach spalteStationsnummer, spalteKartentyp
    
 
    For Each sKanaele In sortKanaele
    
        If sKanaele.Kartentyp.Kartentyp = "CPX 5/2 bistabil" And .Cells(Zaehler1, 1).Value <> 0 Then
        
        


            
            
            
            rData.AddDataSet sKanaele   ' Datensätze von sKanaele in rData schreiben
            
        End If
    Next
    
'Next
    '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
    rData.writeDatsetsToExcel tabelleDaten
    
    
    
    MsgBox "Test"
    
End With

End Sub





