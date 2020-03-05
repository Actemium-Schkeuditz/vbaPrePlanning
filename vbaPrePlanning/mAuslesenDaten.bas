Attribute VB_Name = "mAuslesenDaten"
 Option Explicit

Public Sub AuslesenDaten()

    Dim tabelleDaten As String
    Dim dataKanaele As New cKanalBelegungen
    Dim sKanaele As New cBelegung
    Dim sPerPLCtypKanaele As New cBelegung
    Dim sortKanaele As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen 'neu  CL
    Dim rData As New cKanalBelegungen
    Dim sResult As New cBelegung                 'neu CL
    Dim sKartentyp As New Collection
    Dim karten As Variant
   
    ' Tabellen definieren
    tabelleDaten = "EplSheet"
    
    
    ' Kartentypen definieren
    sKartentyp.Add "CPX 5/2 bistabil"
    sKartentyp.Add "CPX 2x3/2 mono"
    sKartentyp.Add "CPX 5/2 mono"
  
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele 'Auslesen der DAten aus Excelliste
    
    Set sortKanaele = dataKanaele.Sort           'Sortieren nach spalteStationsnummer, spalteKartentyp
    '##### Daten bearbeiten #####
    For Each karten In sKartentyp
    
        Set dataSearchPlcTyp = sortKanaele.searchDatasetPlcTyp(karten) 'neu  CL Suchen nach dem einen Kartentyp
 
        For Each sPerPLCtypKanaele In dataSearchPlcTyp 'neu CL für jeden Datensatz mit dem Kartentyp einmal durchlaufen
            'neu CL schreiben der Daten für das 2.SPS Signal der 5/2 Bistabilen
            If karten = "CPX 5/2 bistabil" Then
                rData.Add sPerPLCtypKanaele.Key, sPerPLCtypKanaele.KWSBMK, 2, sPerPLCtypKanaele.Stationsnummer, vbNullString, sPerPLCtypKanaele.Steckplatz, sPerPLCtypKanaele.Kanal + 1, sPerPLCtypKanaele.Segmentvorlage, sPerPLCtypKanaele.Adress, 0, 0, sPerPLCtypKanaele.SPSBMK
            End If
            
            For Each sKanaele In sortKanaele     'neu CL in allen Kanaele nach den passenden Datensätzen suchen
                If Left(Trim(sPerPLCtypKanaele.KWSBMK), Len(Trim(sPerPLCtypKanaele.KWSBMK)) - 4) = Left(Trim(sKanaele.KWSBMK), Len(Trim(sKanaele.KWSBMK)) - 5) And Right(Trim(sKanaele.KWSBMK), 5) = ".ES01" Then 'suchen nach den Datesätrzen die zusammen gehören Suchen nach .ES01
                
                    'Neu Signal 5
                    If sPerPLCtypKanaele.Signal = 1 Then
                        sResult.Key = sPerPLCtypKanaele.Key
                        sResult.Signal = 5
                        sResult.Steckplatz = sPerPLCtypKanaele.Steckplatz
                        sResult.Kanal = sPerPLCtypKanaele.Kanal
                        sResult.Adress = sPerPLCtypKanaele.Adress
                        sResult.KWSBMK = sPerPLCtypKanaele.KWSBMK
                        sResult.SPSBMK = sPerPLCtypKanaele.SPSBMK
                        sResult.Anschluss1 = sPerPLCtypKanaele.Anschluss1
                        sResult.Anschluss2 = sPerPLCtypKanaele.Anschluss2
                        sResult.Segmentvorlage = sPerPLCtypKanaele.Segmentvorlage
                    
                        rData.AddDataSet sResult ' Datensätze von sKanaele in rData schreiben
                    End If
                    'Neu Signal 6
                    If sPerPLCtypKanaele.Signal = 1 And karten = "CPX 5/2 bistabil" Then
                        sResult.Key = sPerPLCtypKanaele.Key
                        sResult.Signal = 6
                        sResult.Steckplatz = sPerPLCtypKanaele.Steckplatz
                        sResult.Kanal = sPerPLCtypKanaele.Kanal + 1
                        sResult.Adress = Left(Trim(sPerPLCtypKanaele.Adress), Len(Trim(sPerPLCtypKanaele.Adress)) - 1) & CInt(Right(Trim(sPerPLCtypKanaele.Adress), 1)) + 1
                        sResult.KWSBMK = sPerPLCtypKanaele.KWSBMK
                        sResult.SPSBMK = sPerPLCtypKanaele.SPSBMK
                        sResult.Anschluss1 = sPerPLCtypKanaele.Anschluss1
                        sResult.Anschluss2 = sPerPLCtypKanaele.Anschluss2
                        sResult.Segmentvorlage = sPerPLCtypKanaele.Segmentvorlage
                
                        rData.AddDataSet sResult ' Datensätze von sKanaele in rData schreiben
                    End If
                End If
            Next
        Next
    Next
    '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
    rData.writeDatsetsToExcel tabelleDaten
    
    MsgBox "Daten geschrieben"
End Sub











