Attribute VB_Name = "mCPXDatenErgaenzen"
' Skript zur Korrektur der Festo Anschlussdaten
' V0.1
' 09.03.2020
' Christian Langrock
' christian.langrock@actemium.de
' Mohammad Safaadin Hussein
' Mohammad.SafaadinHussein@actemium.de


'@folder (Daten.SPS-Anschlüsse)
 
 Option Explicit

Public Sub CPXDatenErgaenzen()

    Dim tabelleDaten As String
    Dim dataKanaele As New cKanalBelegungen
    Dim sKanaele As New cBelegung
    Dim sPerPLCtypKanaele As New cBelegung
    Dim sortKanaele As New cKanalBelegungen
    Dim dataSearchPlcTyp As New cKanalBelegungen 'neu  CL
    Dim rData As New cKanalBelegungen
    Dim sResult As New cBelegung                 'neu CL
    Dim sKartentyp As New Collection
    Dim Karten As Variant
    Dim bAdressLaenge As Integer                 'benötigte Adresslaenge für Berechnungen
    Dim bLastAdressPos As Integer                'benötigte Letzte Adress-Stelle für Berechnungen
    Dim sPerPLCtypKanaeleAdress2 As String       'Adresse für SPSKanal 2
    Dim iSubAnschluss As Long
   
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    ' Kartentypen definieren
    sKartentyp.Add "CPX 5/2 bistabil"
    sKartentyp.Add "CPX 2x3/2 mono"
    sKartentyp.Add "CPX 5/2 mono"
    
    iSubAnschluss = 0
  
    '##### lesen der belegten Kanäle aus Excel Tabelle #####
    dataKanaele.ReadExcelDataChanelToCollection tabelleDaten, dataKanaele 'Auslesen der DAten aus Excelliste
    
    Set sortKanaele = dataKanaele.Sort           'Sortieren nach spalteStationsnummer, spalteKartentyp
    '##### Daten bearbeiten #####
    For Each Karten In sKartentyp
    
        Set dataSearchPlcTyp = sortKanaele.searchDatasetPlcTyp(Karten) 'neu  CL Suchen nach dem einen Kartentyp
 
        For Each sPerPLCtypKanaele In dataSearchPlcTyp 'neu CL für jeden Datensatz mit dem Kartentyp einmal durchlaufen
            'neu CL schreiben der Daten für das 2.SPS Signal der 5/2 Bistabilen
            If Karten = "CPX 5/2 bistabil" And sPerPLCtypKanaele.Adress <> "" Then
                bAdressLaenge = Len(Trim(sPerPLCtypKanaele.Adress)) - 1 'Adresslaenge - 1 Z.B. "A8503." => 6
                bLastAdressPos = CInt(Right(Trim(sPerPLCtypKanaele.Adress), 1)) + 1 'Letzte Adress-Stelle + 1 z.B. "A8503.0" => 1
                sPerPLCtypKanaeleAdress2 = Left(Trim(sPerPLCtypKanaele.Adress), bAdressLaenge) & bLastAdressPos 'Adresse für SPSKanal 2
                
                rData.Add sPerPLCtypKanaele.Key, sPerPLCtypKanaele.KWSBMK, 2, sPerPLCtypKanaele.Stationsnummer, vbNullString, sPerPLCtypKanaele.Steckplatz, sPerPLCtypKanaele.Kanal + 1, sPerPLCtypKanaele.Segmentvorlage, sPerPLCtypKanaeleAdress2, 0, 0, sPerPLCtypKanaele.SPSBMK
            End If
            
            For Each sKanaele In sortKanaele     'neu CL in allen Kanaele nach den passenden Datensätzen suchen
                If Left(Trim(sPerPLCtypKanaele.KWSBMK), Len(Trim(sPerPLCtypKanaele.KWSBMK)) - 4) = Left(Trim(sKanaele.KWSBMK), Len(Trim(sKanaele.KWSBMK)) - 5) And Right(Trim(sKanaele.KWSBMK), 5) = ".ES01" Then 'suchen nach den Datesätrzen die zusammengehören, Suchen nach .ES01
                       
                    'Neu Signal 5
                    If sPerPLCtypKanaele.Signal = 1 Then
                        sResult.Key = sKanaele.Key
                        sResult.Signal = 5
                        sResult.Steckplatz = sPerPLCtypKanaele.Steckplatz
                        sResult.Kanal = sPerPLCtypKanaele.Kanal
                        sResult.Adress = sPerPLCtypKanaele.Adress
                        sResult.KWSBMK = sPerPLCtypKanaele.KWSBMK
                        sResult.SPSBMK = sPerPLCtypKanaele.SPSBMK
                        sResult.Segmentvorlage = sKanaele.Segmentvorlage
                        ' Anschlüsse schreiben
                        iSubAnschluss = sResult.Kanal Mod 2
                        If iSubAnschluss = 0 Then
                            sResult.Anschluss1 = 2
                        Else
                            sResult.Anschluss1 = 4
                        End If
                        sResult.Anschluss2 = sPerPLCtypKanaele.Anschluss2
                    
                        rData.AddDataSet sResult ' Datensätze von sKanaele in rData schreiben
                    End If
                    
                    'Neu Signal 6
                    If sPerPLCtypKanaele.Signal = 1 And Karten = "CPX 5/2 bistabil" Then
                        sResult.Key = sKanaele.Key
                        sResult.Signal = 6
                        sResult.Steckplatz = sPerPLCtypKanaele.Steckplatz
                        sResult.Kanal = sPerPLCtypKanaele.Kanal + 1
                        sResult.Adress = sPerPLCtypKanaeleAdress2
                        sResult.KWSBMK = sPerPLCtypKanaele.KWSBMK
                        sResult.SPSBMK = sPerPLCtypKanaele.SPSBMK
                        sResult.Segmentvorlage = sKanaele.Segmentvorlage
                        ' Anschlüsse schreiben
                        iSubAnschluss = sResult.Kanal Mod 2
                        If iSubAnschluss = 0 Then
                            sResult.Anschluss1 = 2
                        Else
                            sResult.Anschluss1 = 4
                        End If
                        sResult.Anschluss2 = sPerPLCtypKanaele.Anschluss2
                
                        rData.AddDataSet sResult ' Datensätze von sKanaele in rData schreiben
                    End If
                End If
            Next
        Next
    Next
    '####### Zurückschreiben der Daten in ursprüngliche Excelliste #######
    rData.writeDatsetsToExcel tabelleDaten
    
    'MsgBox "Daten geschrieben"
End Sub
