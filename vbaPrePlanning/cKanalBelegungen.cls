VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKanalBelegungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this collection could contain anything, but the class
'controls access to it and ensures it will just contain Chanels for PLC

'@folder (Daten.Kanalbelegung)

Private objKanalBelegungen As New Collection

Private Sub Class_Initialize()
    Set objKanalBelegungen = New Collection
End Sub

Private Sub Class_Terminate()
    Set objKanalBelegungen = Nothing
End Sub

Public Sub Add(ByRef Key As String, ByRef KWSBMK As String, ByRef Signal As Integer, ByRef Stationsnummer As Integer, ByRef Kartentyp As String, Optional ByRef Steckplatz As Integer, Optional ByRef Kanal As Integer, Optional ByRef Segmentvorlage As String, Optional ByRef Adresse As String, Optional ByRef InputAdressSize As Long, Optional ByRef OutputAdressSize As Long)
    'create a dataset for Chanel and add to collection

    Dim p As New cBelegung
    p.Key = Key
    p.KWSBMK = KWSBMK
    p.Signal = Signal
    p.Stationsnummer = Stationsnummer
    p.Kartentyp.Kartentyp = Kartentyp
    p.Steckplatz = Steckplatz
    p.Kanal = Kanal
    p.Segmentvorlage = Segmentvorlage
    p.Adress = Adresse
    p.InputAdressSize = InputAdressSize
    p.OutputAdressSize = OutputAdressSize
    
    objKanalBelegungen.Add p
End Sub

Public Sub Addobj(ByRef obj As cBelegung)
    objKanalBelegungen.Add obj
End Sub

Public Sub AddDataSet(ByVal dataSet As cBelegung)
    'create a new dataset and add to collection
    Dim p As New cBelegung
    p.Key = dataSet.Key
    p.KWSBMK = dataSet.KWSBMK
    p.Signal = dataSet.Signal
    p.Stationsnummer = dataSet.Stationsnummer
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.Steckplatz = dataSet.Steckplatz
    p.Kanal = dataSet.Kanal
    p.Segmentvorlage = dataSet.Segmentvorlage
    p.Adress = dataSet.Adress
    p.InputAdressSize = dataSet.InputAdressSize
    p.OutputAdressSize = dataSet.OutputAdressSize
    
    objKanalBelegungen.Add p
End Sub

Public Property Get Count() As Long
    'return the number of Datasets
    Count = objKanalBelegungen.Count
End Property

Public Sub Remove(ByRef NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objKanalBelegungen.Remove NameOrNumber
End Sub

'Attribute Item.VB_UserMemId = 0
'@DefaultMember
Public Property Get Item(ByRef Index As Variant) As cBelegung
Attribute Item.VB_UserMemId = 0
    Set Item = objKanalBelegungen.Item(Index)
End Property

'VBA BUG falls die Suchfunktion nicht funktioniert folgendes hier einfügen:
'Attribute NewEnum.VB_UserMemId = -4
'danach im Notepad die exortierte Klasse bearbeiten, die Auskommentierung für: "' Attribute NewEnum.VB_UserMemId = -4" entfernen und wieder importieren
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objKanalBelegungen.[_NewEnum]
End Property

Public Function Sort() As cKanalBelegungen
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim bln As Boolean
    Dim lngCount As Long
    Dim arr() As Long
    Dim ppl As cKanalBelegungen
    
    lngCount = Me.Count
    If lngCount > 0 Then
        ReDim arr(0 To lngCount - 1)
        For i = 0 To lngCount - 1: arr(i) = i + 1: Next

        For i = 1 To lngCount - 1
            k = arr(i)
            j = i - 1
            bln = False
            Do
                If Me.Item(arr(j)).CompareTo(Me.Item(k)) > 0 Then
                    arr(j + 1) = arr(j)
                    j = j - 1
                    If j < 0 Then bln = True
                Else
                    bln = True
                End If
            Loop Until bln
            arr(j + 1) = k
        Next
    End If

    Set ppl = New cKanalBelegungen
    For i = 0 To lngCount - 1: ppl.Addobj Me.Item(arr(i)): Next

    Set Sort = ppl
End Function


Public Function SortSegment() As cKanalBelegungen
    ' Sort Segmentdefinition
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim bln As Boolean
    Dim lngCount As Long
    Dim arr() As Long
    Dim ppl As cKanalBelegungen
    
    lngCount = Me.Count
    If lngCount > 0 Then
        ReDim arr(0 To lngCount - 1)
        For i = 0 To lngCount - 1: arr(i) = i + 1: Next

        For i = 1 To lngCount - 1
            k = arr(i)
            j = i - 1
            bln = False
            Do
                If Me.Item(arr(j)).CompareToSegment(Me.Item(k)) > 0 Then
                    arr(j + 1) = arr(j)
                    j = j - 1
                    If j < 0 Then bln = True
                Else
                    bln = True
                End If
            Loop Until bln
            arr(j + 1) = k
        Next
    End If

    Set ppl = New cKanalBelegungen
    For i = 0 To lngCount - 1: ppl.Addobj Me.Item(arr(i)): Next

    Set SortSegment = ppl
End Function

Public Function SortChannel() As cKanalBelegungen
    ' Sort Station, Slot and Channel
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim bln As Boolean
    Dim lngCount As Long
    Dim arr() As Long
    Dim ppl As cKanalBelegungen
    
    lngCount = Me.Count
    If lngCount > 0 Then
        ReDim arr(0 To lngCount - 1)
        For i = 0 To lngCount - 1: arr(i) = i + 1: Next

        For i = 1 To lngCount - 1
            k = arr(i)
            j = i - 1
            bln = False
            Do
                If Me.Item(arr(j)).CompareToChannel(Me.Item(k)) > 0 Then
                    arr(j + 1) = arr(j)
                    j = j - 1
                    If j < 0 Then bln = True
                Else
                    bln = True
                End If
            Loop Until bln
            arr(j + 1) = k
        Next
    End If

    Set ppl = New cKanalBelegungen
    For i = 0 To lngCount - 1: ppl.Addobj Me.Item(arr(i)): Next

    Set SortChannel = ppl
End Function



Public Sub ReadExcelDataChanelToCollection(ByRef tablename As String, ByRef dataColection As cKanalBelegungen, ByRef spalteStationsnummer As String, ByRef spalteKartentyp As String)
    'Reads the table on the "EplSheet" sheet and adds
    'Chanels to the collection. It also adds values
    'to the product's properties.
    Dim rCell As Range
    Dim rTable As Range
    Dim sSpalteEingangsAdress As String
    Dim sSpalteAusgangsAdress As String
    Dim iSpalteEingangsAdress As Long
    Dim iSpalteAusgangsAdress As Long
    Dim iSpalteIntStationsnummer As Long
    Dim ispalteKartentyp As Long
    
    On Error GoTo ErrorHandle

    sSpalteEingangsAdress = "BD"
    sSpalteAusgangsAdress = "BE"
    ' Konvertierung Spaltenbuchstaben in LONG
    iSpalteIntStationsnummer = SpaltenBuchstaben2Int(spalteStationsnummer) - 1
    iSpalteEingangsAdress = SpaltenBuchstaben2Int(sSpalteEingangsAdress) - 1
    iSpalteAusgangsAdress = SpaltenBuchstaben2Int(sSpalteAusgangsAdress) - 1
       
    ispalteKartentyp = SpaltenBuchstaben2Int(spalteKartentyp) - 1
    'First we check if the table is intact. The
    'check is NOT complete.
    ThisWorkbook.Worksheets(tablename).Activate
    'Filter aus, aber nicht löschen
            If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    If Len(Range("A3")) = 0 Then
        MsgBox "The table must start in cell A3. Adding products has been aborted.", vbCritical, "Error"
        Set dataColection = Nothing
        Exit Sub
    End If

    Set rTable = Range("A3")

    With rTable
        If Len(.Offset(0, 0).Value) = 0 Or Len(.Offset(1, 0).Value) = 0 Then
            MsgBox "The products table has been destoyed and there is no or too few data.", vbCritical, "Error"
            Set dataColection = Nothing
            Exit Sub
        End If
    End With

    'If the table is okay, we start reading data.
    'The range is expanded to include the last
    'cell with content.
    If Len(rTable.Offset(2, 0).Formula) = 0 Then
        Set rTable = rTable.Offset(0, 9)
    Else
        Set rTable = Range(rTable.Offset(0, 0), rTable.End(xlDown))
    End If

    For Each rCell In rTable
        With rCell
            ' prüfen ob Stationsnummer nicht leer
            If rCell.Offset(0, iSpalteIntStationsnummer).Value <> vbNullString Then
                '      sEinbauort = str$(.Value)
                ' die gelesenen Daten anhängen
                'Signal 1
                 If rCell.Offset(0, ispalteKartentyp).Value <> vbNullString Then
                Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 1, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp).Value, CInt(rCell.Offset(0, ispalteKartentyp + 2).Value), CInt(rCell.Offset(0, ispalteKartentyp + 3).Value), rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4).Value, CLng(rCell.Offset(0, iSpalteEingangsAdress)), CLng(rCell.Offset(0, iSpalteAusgangsAdress))
              End If
                'Signal 2
                If rCell.Offset(0, ispalteKartentyp + 12).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 2, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 12).Value, rCell.Offset(0, ispalteKartentyp + 2 + 12).Value, rCell.Offset(0, ispalteKartentyp + 3 + 12).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 12).Value ', CLng(rCell.Offset(0, iSpalteEingangsAdress)), CLng(rCell.Offset(0, iSpalteAusgangsAdress))
                End If
                'Signal 3
                If rCell.Offset(0, ispalteKartentyp + 24).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 3, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 24).Value, rCell.Offset(0, ispalteKartentyp + 2 + 24).Value, rCell.Offset(0, ispalteKartentyp + 3 + 24).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 24).Value
                End If
                'Signal 4
                If rCell.Offset(0, ispalteKartentyp + 36).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 4, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 36).Value, rCell.Offset(0, ispalteKartentyp + 2 + 36).Value, rCell.Offset(0, ispalteKartentyp + 3 + 36).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 36).Value
                End If
                'Signal 5
                If rCell.Offset(0, ispalteKartentyp + 48).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 5, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 48).Value, rCell.Offset(0, ispalteKartentyp + 2 + 48).Value, rCell.Offset(0, ispalteKartentyp + 3 + 48).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 48).Value
                End If
            End If
        End With
    Next
    'MsgBox "gelesen"

BeforeExit:
    Set rCell = Nothing
    Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim einlesen.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Sub writeDatsetsToExcel(ByRef tablename As String)
    Dim sdata As New cBelegung
    Dim ws1 As Worksheet
    Dim sSpalteSteckplatz As String
    Dim sSpalteKanal As String
    Dim sSpalteSegmentvorlage As String
    Dim iSpalteSteckplatz As Long
    Dim iSpalteKanal As Long
    Dim i As Long
    Dim zeilenanzahl As Long
        
    ' Spalten definieren
    sSpalteSteckplatz = "CA"
    sSpalteKanal = "CB"
    sSpalteSegmentvorlage = "BL"
    ' nummerischen Wert ermitteln
    iSpalteSteckplatz = SpaltenBuchstaben2Int(sSpalteSteckplatz)
    iSpalteKanal = SpaltenBuchstaben2Int(sSpalteKanal)
        
        
    ThisWorkbook.Worksheets(tablename).Activate
    'Set wkb = ActiveWorkbook
    'Set ws1 = Worksheets.[_Default](tabelleDaten)
    Set ws1 = ThisWorkbook.ActiveSheet
    ' Tabelle mit Daten bearbeiten
    With ws1
        
        ' Worksheet activate
        ws1.Activate

        Application.ScreenUpdating = False
        '##### Herausfinden der Anzahl der Zeilen #####
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt

        
        For Each sdata In Me
            For i = 3 To zeilenanzahl
                ' suchen nach dem pasenden key
                If .Cells.Item(i, "A") = sdata.Key Then
                    .Cells.Item(i, (iSpalteSteckplatz + (12 * (sdata.Signal - 1)))) = sdata.Steckplatz
                    .Cells.Item(i, (iSpalteKanal + (12 * (sdata.Signal - 1)))) = sdata.Kanal
                    .Cells.Item(i, (iSpalteKanal + 1 + (12 * (sdata.Signal - 1)))) = sdata.Adress
                    .Cells.Item(i, sSpalteSegmentvorlage) = sdata.Segmentvorlage
                    'MsgBox "Kanal geschrieben: " & sData.Kanal
                End If
            
            Next i
        Next
    End With

BeforeExit:
    'Set rCell = Nothing
    'Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim einlesen.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Function searchKanalBelegungenKartentyp(ByRef pStationsnummer As Collection, ByVal pKartentyp As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Kartentyp und sortiert nach Stationsnummer
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
    Set searchKanalBelegungenKartentyp = Nothing
    Dim Station As Variant
    
    For Each Station In pStationsnummer
        For Each sdata In Me
            If sdata.Stationsnummer = Station Then
                If sdata.Kartentyp.Kartentyp = pKartentyp Then
                    rData.AddDataSet sdata
                    'Debug.Print searchAnschluss.Item(1).Kartentyp;
                    '    Exit For                         ' nur einmal suchen dann beenden
                End If
            End If
        Next
    Next
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchKanalBelegungenKartentyp = rData
    'Test der Suche
    '  Debug.Print searchAnschluss.Item(1).Kartentyp; vbTab; searchAnschluss.Item(1).Kanal; vbTab; searchAnschluss.Item(1).Anschluss_1; vbTab; searchAnschluss.Item(1).Anschluss_2; vbTab; searchAnschluss.Item(1).Anschluss_3; vbTab; searchAnschluss.Item(1).Anschluss_4; vbTab; searchAnschluss.Item(1).Anschluss_M; vbTab; searchAnschluss.Item(1).Anschluss_VS
End Function

Public Function searchKanalBelegungenSteckplatz(ByVal pStationsnummer As Long, ByVal pSteckplatz As Integer) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Steckplatz und sortiert nach Stationsnummer
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
    Set searchKanalBelegungenSteckplatz = Nothing
    'Dim Station As Variant
    
    'For Each Station In pStationsnummer
        For Each sdata In Me
            If sdata.Stationsnummer = pStationsnummer Then
                If sdata.Steckplatz = pSteckplatz Then
                    rData.AddDataSet sdata
                    'Debug.Print searchAnschluss.Item(1).Kartentyp;
                    '    Exit For                         ' nur einmal suchen dann beenden
                End If
            End If
        Next
    'Next
    If rData Is Nothing Then
    If rData.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    End If
    
    Set searchKanalBelegungenSteckplatz = rData
    'Test der Suche
    '  Debug.Print searchAnschluss.Item(1).Kartentyp; vbTab; searchAnschluss.Item(1).Kanal; vbTab; searchAnschluss.Item(1).Anschluss_1; vbTab; searchAnschluss.Item(1).Anschluss_2; vbTab; searchAnschluss.Item(1).Anschluss_3; vbTab; searchAnschluss.Item(1).Anschluss_4; vbTab; searchAnschluss.Item(1).Anschluss_M; vbTab; searchAnschluss.Item(1).Anschluss_VS
End Function

Public Function searchDatasetPerStation(ByVal pStationsnummer As Integer) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einer bestimmten  Stationsnummer
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sdata In Me
        If sdata.Stationsnummer = pStationsnummer Then
            rData.AddDataSet sdata
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetPerStation = rData

End Function

Public Function searchDatasetPlcTyp(ByVal pKartentyp As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Kartentyp
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sdata In Me
        If sdata.Kartentyp.Kartentyp = pKartentyp Then
            rData.AddDataSet sdata
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetPlcTyp = rData
   
End Function

Public Function zuweisenKanal(OffsetSlot As Integer, ByVal pKartentyp As String, ByVal dataPLCConfig As cPLCconfig) As cKanalBelegungen

    ' hochzählen der Kanäle und zurückgeben als Class
    Dim iKanal As Integer
    Dim iOldKanal As Integer
    Dim iSteckplatz As Integer

    Dim rData As New cKanalBelegungen
    Dim sortData As New cKanalBelegungen
    Dim sKWSBMKold As String
    Dim bET200SPNothalt As Boolean
    Dim bRunOnce As Boolean
    Dim sSegmentcorrection As String
    Dim iReserveChannelPerSlot As Long
    Dim iReserveChannelsBefor As Long
    Dim iReserveChannelsAfter As Long
    Dim i As Long

    
    i = 1
    iReserveChannelPerSlot = dataPLCConfig.Item(i).ReserveChannelPerSlot
    iReserveChannelsBefor = dataPLCConfig.Item(i).ReserveChannelsBefor
    iReserveChannelsAfter = dataPLCConfig.Item(i).ReserveChannelsAfter
    
    
    ' Startwerte
    iKanal = 0
    bET200SPNothalt = False
    bRunOnce = False
   iOldKanal = 0
    If OffsetSlot = 0 Then
        iSteckplatz = 1
    Else
        iSteckplatz = 0
    End If
    
    sKWSBMKold = vbNullString

    Dim sdata As New cBelegung
    If pKartentyp = "ET200SP 8FDI" Then
     'erstmal die Kanäle davor bearbeiten
        Do While iReserveChannelsBefor > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                rData.Add "LEER", "LEER", 0, sdata.Stationsnummer, sdata.Kartentyp, iSteckplatz, iKanal
                iSteckplatz = iSteckplatz + 1
                
            End If
            iReserveChannelsBefor = iReserveChannelsBefor - 1
        Loop
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sdata In sortData
            ' Kanäle müssen passend aufgelgt werden
            If Left(sdata.Segmentvorlage, 20) = "FDI_ET200SP_NOT_HALT" Or sdata.Segmentvorlage = "FDI_ET200SP_BMA" Or ((sdata.Segmentvorlage = "FDI_ET200SP_ZUHALTUNG_AZM201B") And (sdata.Signal = 1 Or sdata.Signal = 2)) Then
                bET200SPNothalt = True           ' Hier FDI doppelte Auswertung der Signale
                If sdata.Signal = 2 Then
                    ' Kanäle 4, 5, 6 und 7
                    sdata.Kanal = iOldKanal + 4
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    If iKanal <= ((sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) / 2) Then
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + OffsetSlot
                        iOldKanal = iKanal
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iOldKanal = iKanal
                        iSteckplatz = iSteckplatz + 1
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + OffsetSlot
                        iKanal = iKanal + 1
                      
                    End If
                End If
            ElseIf bET200SPNothalt = True Then
            ' hier weiter wenn einmal ein Not-Aus kam
                If bRunOnce = False Then
                    ' Kanäle 4, 5, 6 und 7
                    sdata.Kanal = iOldKanal + 4
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    bRunOnce = False
                    If iKanal <= ((sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) / 2) Then
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + OffsetSlot
                        iOldKanal = iKanal
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iOldKanal = iKanal
                        iSteckplatz = iSteckplatz + 1
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + OffsetSlot
                         iKanal = iKanal + 1
                    End If
                End If
                'normale Belegung
            Else
                sdata.Kanal = iKanal
                sdata.Steckplatz = iSteckplatz + OffsetSlot
                If iKanal < sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                    iKanal = iKanal + 1
                Else
                    iKanal = 0
                     rData.Add "LEER", "LEER", 0, sdata.Stationsnummer, sdata.Kartentyp, iSteckplatz, iKanal
                    iSteckplatz = iSteckplatz + 1
                End If
            End If
            sKWSBMKold = sdata.KWSBMK
            rData.AddDataSet sdata
        Next
        'die  nachfolgenden Kanäle bearbeiten
        Do While iReserveChannelsAfter > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                 rData.Add "LEER", "LEER", 0, sdata.Stationsnummer, sdata.Kartentyp, iSteckplatz, iKanal
            End If
            iReserveChannelsAfter = iReserveChannelsAfter - 1
             rData.Add "RESERVE", "RESERVE", "1", sdata.Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
        Loop
        
    ElseIf pKartentyp = "CPX-8DE-D" Then
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sdata In sortData
            'pro Stecker werden zwei Kanäle belegt
            sSegmentcorrection = Right(sdata.Segmentvorlage, 2)
            If iKanal = 1 Or iKanal = 3 Or iKanal = 5 Or iKanal = 7 Then
                'ungerade Kanäle 1, 3, 5 und 7
                If iKanal <= (sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) Then
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                  
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                End If
                'sData.Segmentvorlage = "ungerade"
                If sSegmentcorrection = "_1" Then
               sdata.Segmentvorlage = Left(sdata.Segmentvorlage, Len(sdata.Segmentvorlage) - 2) & "_2"
               End If
            Else
                ' Kanäle 0, 2, 4 und 6
                If iKanal < (sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) Then
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                    iOldKanal = iKanal + 4
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + OffsetSlot
                End If
               If sSegmentcorrection = "_2" Then
               sdata.Segmentvorlage = Left(sdata.Segmentvorlage, Len(sdata.Segmentvorlage) - 2) & "_1"
               End If
       
            End If
            rData.AddDataSet sdata
            iKanal = iKanal + 1
        Next
    Else
        'hier bearbeiten alles was keine extra Behandlung benötigt
        'erstmal die Kanäle davor bearbeiten
        Do While iReserveChannelsBefor > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                 rData.Add "LEER", "LEER", 0, sdata.Stationsnummer, sdata.Kartentyp, iSteckplatz, iKanal
                iSteckplatz = iSteckplatz + 1
            End If
            iReserveChannelsBefor = iReserveChannelsBefor - 1
        Loop
        
        For Each sdata In Me
       
            sdata.Kanal = iKanal
            sdata.Steckplatz = iSteckplatz + OffsetSlot
            rData.AddDataSet sdata
            If iKanal < sdata.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
            End If
        Next
        'die  nachfolgenden Kanäle bearbeiten
        Do While iReserveChannelsAfter > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                 rData.Add "LEER", "LEER", 0, sdata.Stationsnummer, sdata.Kartentyp, iSteckplatz, iKanal
            End If
            iReserveChannelsAfter = iReserveChannelsAfter - 1
        Loop
        
    End If
    Set zuweisenKanal = rData
End Function

Public Function returnStation() As Collection
    ' collect all Stations without duplicates
    On Error GoTo ErrorHandler
    Set returnStation = Nothing

    Dim col As New Collection
    Dim bSearchinCol As Boolean
    Dim it As Variant
    Dim sdata As New cBelegung
    
    For Each sdata In Me
        ' prüfen ob Stationsnummer schon in Collection
        bSearchinCol = True
        For Each it In col
            If it = sdata.Stationsnummer Then
                bSearchinCol = False
                
            End If
        Next
        If bSearchinCol = True Then
            If sdata.Stationsnummer <> 0 Then
                col.Add sdata.Stationsnummer     '  dynamically add value to the end
            End If
        End If
    Next
    ' Rückgabe der Daten
    Set returnStation = col

    Exit Function
ErrorHandler:
    If Err.Number > 0 Then
        MsgBox "Fehler in den Daten"
        Err.Clear
        Resume Next
    End If
End Function

Public Function returnKartentyp() As Collection

    ' collect all Kartentyp without duplicates
    On Error GoTo ErrorHandler
    Set returnKartentyp = Nothing

    Dim col As New Collection
    Dim bSearchinCol As Boolean
    Dim it As Variant
    Dim sdata As New cBelegung
    
    For Each sdata In Me
        ' prüfen ob Kartentyp schon in Collection
        bSearchinCol = True
        For Each it In col
            If it = sdata.Kartentyp.Kartentyp Then
                bSearchinCol = False
                
            End If
        Next
        If bSearchinCol = True Then
            If sdata.Kartentyp.Kartentyp <> vbNullString Then
                col.Add sdata.Kartentyp.Kartentyp          '  dynamically add value to the end
            End If
        End If
    Next
    ' Rückgabe der Daten
    Set returnKartentyp = col

    Exit Function
ErrorHandler:
    If Err.Number > 0 Then
        MsgBox "Fehler in den Daten"
        Err.Clear
        Resume Next
    End If
End Function

Public Function returnLastSlotNumber() As Integer

    Dim maxSlotNumber As Integer
    Dim sdata As New cBelegung
    maxSlotNumber = 0

    For Each sdata In Me
        If sdata.Steckplatz > maxSlotNumber Then
        maxSlotNumber = sdata.Steckplatz
        End If
    Next
    returnLastSlotNumber = maxSlotNumber
End Function

Public Function returnFirstSlotNumber() As Integer

    Dim minSlotNumber As Integer
    Dim sdata As New cBelegung
   
    minSlotNumber = 1000

    For Each sdata In Me
        If minSlotNumber > sdata.Steckplatz Then
        minSlotNumber = sdata.Steckplatz
        End If
    
    Next
    returnFirstSlotNumber = minSlotNumber
End Function

Public Function returnAllSlotsPerRack() As cKanalBelegungen
    ' Rückgabe der Steckplätze mit dem ersten Datensatz, bzw Kanal 0
    ' todo not testet
    ' On Error GoTo ErrorHandler
    Set returnAllSlotsPerRack = Nothing
    Dim icount As Long
    '  Dim sData As New cBelegung
    'Dim col As New cKanalBelegungen
    ' Dim col As Variant
    
    ' Suchen nach allen Datensätzten dem ersten belegten Kanal
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
    Set rData = Nothing
   
    For Each sdata In Me
        If rData Is Nothing Then
            icount = 0
        Else
            icount = rData.Count
        End If
    
        If sdata.Kanal = 0 Then
            rData.AddDataSet sdata
        ElseIf sdata.Kanal = 1 And Not (rData Is Nothing) And icount >= 1 Then
            If rData.Item(icount).Steckplatz <> sdata.Steckplatz Then
                rData.AddDataSet sdata
            End If
        Else
            ' MsgBox " Karte ohne Belegung, Stationsnummer" & sData.Stationsnummer & "Steckplatz: " & sData.Steckplatz
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set returnAllSlotsPerRack = rData
      
    ' Dim bSearchinCol As Boolean
    'Dim it As Variant
    '  Dim i As Long
    '   Dim iItems As Long
    
    ' suche ersten Datensatz

    'iItems = col.Count
     
    ' For Each sData In Me
    ' prüfen ob Steckplatz schon in Collection
    '    bSearchinCol = True
    
    '   If iItems > 0 Then
    '       For i = 0 To iItems
    '           If col.Item(i).Steckplatz = sData.Steckplatz Then
    '               bSearchinCol = False
    '           End If
          
    '   Next
            
    '   End If
    '   If bSearchinCol = True Then
    '       If sData.Steckplatz <> vbNullString Then
    '           col.Addobj sData                 '  dynamically add value to the end
    '      End If
    '  End If
    'Next
    ' Rückgabe der Daten
    'Set returnAllSlotsPerRack = col

    Exit Function
ErrorHandler:
    If Err.Number > 0 Then
        MsgBox "Steckplätze konnten nicht ermittelt werden "
        Err.Clear
        Resume Next
    End If
End Function

Public Function sumAdresses(ByVal Station As Integer, ByVal dataPLCConfig As cPLCconfig) As cKanalBelegungen
    'berechnen pro Station
    'sortiert nach Steckplatz und Kanal
    Dim InputAdress As Long
    Dim OutputAdress As Long
    Dim iStartInputAdress As Long
    Dim iStartOutputAdress As Long
    Dim iTmpInputAdress As Long
    Dim iTmpOutputAdress As Long
    Dim iStartSteckplatz As Long
    Dim iSteckplatzdifferenz As Long
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
    Dim sPLCData As New cPLCconfigData
    Dim sPLCDataAdress As New cPLCconfig
    Dim sPLCDataAdressPerSlot As New cPLCconfigData
    Dim sPLCassignmentsPerSlot As New cKanalBelegungen
    Dim sDataPLCSlot As New cBelegung
      
    ' Startwerte setzen
    iStartInputAdress = dataPLCConfig.Item(1).FirstInputAdress
    iStartOutputAdress = dataPLCConfig.Item(1).FirstOutputAdress
    iStartSteckplatz = dataPLCConfig.Item(1).Steckplatz
    
   
    
    
    'Sortieren nach Steckplatz und Kanal
    Me.SortChannel
    ' Startwerte setzen
    iTmpInputAdress = 0
    iTmpOutputAdress = 0
    'Schreiben der Adressen an die Signale
    For Each sdata In Me
        ' auslesen der ermittelten Adressen
        For Each sPLCDataAdressPerSlot In sPLCDataAdress
            If sPLCDataAdressPerSlot.Stationsnummer = sdata.Stationsnummer And sPLCDataAdressPerSlot.Steckplatz = sdata.Steckplatz Then
                InputAdress = sPLCDataAdressPerSlot.FirstInputAdress
                OutputAdress = sPLCDataAdressPerSlot.FirstOutputAdress
            End If
        Next
        
        If sdata.Kartentyp.VariableAdressRange = False Then
            If sdata.Kartentyp.Datatyp = "Bit" Then
                If sdata.Kanal >= 8 Then
                    ' Berechnung wenn Kanal größer 7
                    ' -> +1 Byte
                    If sdata.Kartentyp.InputAdressLength > 0 Then
                        sdata.Adress = "E" & InputAdress + 1 & "." & sdata.Kanal - 8
                    End If
                    If sdata.Kartentyp.OutputAdressLength > 0 Then
                        sdata.Adress = "A" & OutputAdress + 1 & "." & sdata.Kanal - 8
                    End If
                Else
                    If sdata.Kartentyp.InputAdressLength > 0 Then
                        sdata.Adress = "E" & InputAdress & "." & sdata.Kanal
                    End If
                    If sdata.Kartentyp.OutputAdressLength > 0 Then
                        sdata.Adress = "A" & OutputAdress & "." & sdata.Kanal
                    End If
                End If
            ElseIf sdata.Kartentyp.Datatyp = "Byte" Then
                'Ausgabe Byte
                If sdata.Kartentyp.InputAdressLength > 0 Then
                    sdata.Adress = "EB" & InputAdress + (sdata.Kanal * sdata.Kartentyp.InputAdressLength)
                End If
                If sdata.Kartentyp.OutputAdressLength > 0 Then
                    sdata.Adress = "AB" & OutputAdress + (sdata.Kanal * sdata.Kartentyp.OutputAdressLength)
                End If
                 ElseIf sdata.Kartentyp.Datatyp = "Word" Then
                'Ausgabe Word
                If sdata.Kartentyp.InputAdressLength > 0 Then
                    sdata.Adress = "EW" & InputAdress + (sdata.Kanal * sdata.Kartentyp.InputAdressLength)
                End If
                If sdata.Kartentyp.OutputAdressLength > 0 Then
                    sdata.Adress = "AW" & OutputAdress + (sdata.Kanal * sdata.Kartentyp.OutputAdressLength)
                End If
            End If
        Else
            If sPLCData.Steckplatz > iStartSteckplatz Then
                ' erkenne wenn sich Steckplatz ändert
                iStartSteckplatz = sdata.Steckplatz
                iTmpInputAdress = 0
                iTmpOutputAdress = 0
            End If
            'variable Adressvergabe
            'Ausgabe Byte
            If sdata.Kartentyp.Datatyp = "Byte" Then
                If sdata.InputAdressSize > 0 Then
                    sdata.Adress = "EW" & InputAdress + iTmpInputAdress
                    iTmpInputAdress = iTmpInputAdress + sdata.InputAdressSize
                End If
                If sdata.OutputAdressSize > 0 Then
                    sdata.Adress = "AW" & OutputAdress + iTmpOutputAdress
                    iTmpOutputAdress = iTmpOutputAdress + sdata.OutputAdressSize
                End If
            Else
                MsgBox "Anderer Datentyp als 'Byte' ist nicht vorgesehen für Datensatz: " & sdata.KWSBMK
            End If
             
        End If
        
    Next
    
    
    Set sumAdresses = rData

End Function



Public Function AdressPerSlottyp(ByRef iInputStartAdress As Long, ByRef iOutputStartAdress As Long, ByVal iStationsnummer As Integer, ByVal pKartentyp As String) As cKanalBelegungen
    'Berechnen der Adressen, wegen der Variabilität der IO-Link Geräte nur möglich wenn die gesamte Station bzw. der Kartentyp belegt ist
    
    Dim iTmpAdressInput As Long
    Dim iTmpAdressOutput As Long
    Dim iAdressVariInput As Long
    Dim iAdressVariOutput As Long
    Dim ifirstSlot As Integer
    Dim ilastSlot As Integer
    Dim sdata As New cBelegung
    Dim sdataVari As New cBelegung
    Dim SlotData As New cKanalBelegungen
    Dim rData As New cKanalBelegungen
    Dim i As Long
    Dim y As Long
    Dim x As Long
    
    ' sortieren nach Station, Slot and Channel
    Me.SortChannel
    'Me.returnAllSlotsPerRack
   
    ' Steckplätze ermitteln
    ifirstSlot = Me.returnFirstSlotNumber
    ilastSlot = Me.returnLastSlotNumber
    
    If Me.Item(1).Kartentyp.VariableAdressRange = True Then ' in einem Datensatz kommt nur ein Kartentyp vor
      
        ' ermiteln für jede Karte die Belegung
        For i = ifirstSlot To ilastSlot
            'fixed adress range for input and output
            ' prüfe was größer ist, diesen Wert nehmen
            If Me.Item(1).Kartentyp.SameAdressRangeForOutput = True Then
                If iInputStartAdress > iOutputStartAdress Then
                    iOutputStartAdress = iInputStartAdress
                Else
                    iInputStartAdress = iOutputStartAdress
                End If
            End If
            Set SlotData = Nothing
            ' Durchlauf für jeden Steckplatz einzeln
            Set SlotData = Me.searchKanalBelegungenSteckplatz(iStationsnummer, i)
            iAdressVariInput = 0
            iAdressVariOutput = 0
            
            If pKartentyp = "ET200SP 4IO-LINK" Then ' bei ET200SP IO-Link alle Kanäle zusammen zählen
                For Each sdataVari In SlotData
                    'hier Adresse ermitteln
                    sdataVari.Adress = "EB" & iInputStartAdress + iAdressVariInput ' + sdataVari.InputAdressSize
                    ' Adressen hochzählen
                    iAdressVariInput = iAdressVariInput + sdataVari.InputAdressSize
                    iAdressVariOutput = iAdressVariOutput + sdataVari.OutputAdressSize
                    sdataVari.InputAdressPerSlot = iInputStartAdress
                    sdataVari.OutputAdressPerSlot = iOutputStartAdress
                   
                    ' Ergebnis einsammeln
                    rData.Addobj sdataVari
                Next
                'Berechnen Adressen für die nächste Karte
                If iAdressVariInput = 1 And iAdressVariOutput <= 1 Then
                    iInputStartAdress = iInputStartAdress + 1
                    iOutputStartAdress = iOutputStartAdress + 1
                ElseIf iAdressVariInput <= 1 And iAdressVariOutput = 1 Then 'beide null ist Fehler
                    iInputStartAdress = iInputStartAdress + 1
                    iOutputStartAdress = iOutputStartAdress + 1
                ElseIf iAdressVariInput <= 2 And iAdressVariOutput <= 2 Then
                    iInputStartAdress = iInputStartAdress + 2
                    iOutputStartAdress = iOutputStartAdress + 2
                ElseIf iAdressVariInput <= 8 And iAdressVariOutput <= 8 Then
                    iInputStartAdress = iInputStartAdress + 8
                    iOutputStartAdress = iOutputStartAdress + 8
                ElseIf iAdressVariInput <= 16 And iAdressVariOutput <= 16 Then
                    iInputStartAdress = iInputStartAdress + 16
                    iOutputStartAdress = iOutputStartAdress + 16
                ElseIf iAdressVariInput <= 32 And iAdressVariOutput <= 32 Then
                    iInputStartAdress = iInputStartAdress + 32
                    iOutputStartAdress = iOutputStartAdress + 32
                ElseIf iAdressVariInput <= 64 And iAdressVariOutput <= 64 Then
                    iInputStartAdress = iInputStartAdress + 64
                    iOutputStartAdress = iOutputStartAdress + 64
                ElseIf iAdressVariInput <= 128 And iAdressVariOutput <= 128 Then
                    iInputStartAdress = iInputStartAdress + 144
                    iOutputStartAdress = iOutputStartAdress + 128
                Else
                    MsgBox "Fehler in der Adresslänge von: " & sdata.KWSBMK
                End If
                'iAdressVariInput = 0
                'iAdressVariOutput = 0
                iAdressVariInput = iInputStartAdress
                    iAdressVariOutput = iOutputStartAdress
            ElseIf pKartentyp = "AL1400" Or pKartentyp = "AL1402" Then 'Berechnung pro Kanal und dann zusammen zählen
                For Each sdataVari In SlotData
                    
                    If sdataVari.Kanal = 0 Then
                        iAdressVariInput = iAdressVariInput + sdataVari.Kartentyp.InputAdressDiagnosticLength 'Wenn Kanal 0 dann Diagnosebereich belegen
                        iAdressVariOutput = iAdressVariOutput + sdataVari.Kartentyp.OutputAdressDiagnosticLength
                    End If
                    sdataVari.InputAdressPerSlot = iInputStartAdress
                    sdataVari.OutputAdressPerSlot = iOutputStartAdress
                    
                    If sdataVari.InputAdressSize = 1 And sdataVari.OutputAdressSize <= 1 Then
                        iAdressVariInput = iAdressVariInput + 1
                        iAdressVariOutput = iAdressVariOutput + 1
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 1 And sdataVari.OutputAdressSize = 1 Then 'beide null ist Fehler
                        iAdressVariInput = iAdressVariInput + 1
                        iAdressVariOutput = iAdressVariOutput + 1
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 2 And sdataVari.OutputAdressSize <= 2 Then
                        iAdressVariInput = iAdressVariInput + 2
                        iAdressVariOutput = iAdressVariOutput + 2
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 4 And sdataVari.OutputAdressSize <= 4 Then
                        iAdressVariInput = iAdressVariInput + 4
                        iAdressVariOutput = iAdressVariOutput + 4
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 8 And sdataVari.OutputAdressSize <= 8 Then
                        iAdressVariInput = iAdressVariInput + 8
                        iAdressVariOutput = iAdressVariOutput + 8
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 16 And sdataVari.OutputAdressSize <= 16 Then
                        iAdressVariInput = iAdressVariInput + 16
                        iAdressVariOutput = iAdressVariOutput + 16
                        sdataVari.Adress = "EB" & iAdressVariInput
                    ElseIf sdataVari.InputAdressSize <= 32 And sdataVari.OutputAdressSize <= 32 Then
                        iAdressVariInput = iAdressVariInput + 32
                        iAdressVariOutput = iAdressVariOutput + 32
                        sdataVari.Adress = "EB" & iAdressVariInput
                    Else
                        MsgBox "Fehler in der Adresslänge von: " & sdata.KWSBMK
                    End If
                    
                    rData.Addobj sdataVari
                Next
                ' Adressen für nächste Karte schreiben
                iInputStartAdress = iAdressVariInput
                iOutputStartAdress = iAdressVariOutput
                iAdressVariInput = 0
                iAdressVariOutput = 0
            Else
                MsgBox "Adressen für Kartentyp: " & sdataVari.Kartentyp.Kartentyp & " kann nicht berechnet werden"
            End If
        Next i
    Else
        ' hier die anderen Karten
        'fixed adress range for input and output
        ' prüfe was größer ist, diesen Wert nehmen
            If Me.Item(1).Kartentyp.SameAdressRangeForOutput = True Then
                If iInputStartAdress > iOutputStartAdress Then
                    iOutputStartAdress = iInputStartAdress
                Else
                    iInputStartAdress = iOutputStartAdress
                End If
            End If
        ' ermiteln für jede Karte einzeln die Belegung
        For y = ifirstSlot To ilastSlot
            Set SlotData = Nothing
            ' Durchlauf für jeden Steckplatz einzeln
            Set SlotData = Me.searchKanalBelegungenSteckplatz(iStationsnummer, i)
            iAdressVariInput = 0
            iAdressVariOutput = 0
        
            For Each sdata In Me
                sdata.InputAdressPerSlot = iInputStartAdress
                sdata.OutputAdressPerSlot = iOutputStartAdress
                If sdata.Kartentyp.SameAdressRangeForOutput Then
                    ' prüfen was größer ist
                    If iInputStartAdress > iOutputStartAdress Then
                        sdata.InputAdressPerSlot = iInputStartAdress
                        sdata.OutputAdressPerSlot = sdata.InputAdressPerSlot
                    Else
                        sdata.OutputAdressPerSlot = iOutputStartAdress
                        sdata.InputAdressPerSlot = sdata.OutputAdressPerSlot
                    End If
                Else
                    ' hier normale Karten ohne extras
                    sdata.InputAdressPerSlot = iInputStartAdress
                    sdata.OutputAdressPerSlot = iOutputStartAdress
                End If
                If sdata.Kartentyp.Datatyp = "Byte" Then
                    If sdata.Kartentyp.SignalDirection = "Input" Or sdata.Kartentyp.SignalDirection = "InOut" Then
                        sdata.Adress = "EB" & iInputStartAdress + (sdata.Kanal * sdata.Kartentyp.InputAdressLength)
                    ElseIf sdata.Kartentyp.SignalDirection = "Output" Then
                        sdata.Adress = "AB" & iOutputStartAdress + (sdata.Kanal * sdata.Kartentyp.OutputAdressLength)
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sdata.Kartentyp.Kartentyp
                    End If
                ElseIf sdata.Kartentyp.Datatyp = "Word" Then
                    If sdata.Kartentyp.SignalDirection = "Input" Or sdata.Kartentyp.SignalDirection = "InOut" Then
                        sdata.Adress = "EW" & iInputStartAdress + (sdata.Kanal * sdata.Kartentyp.InputAdressLength)
                    ElseIf sdata.Kartentyp.SignalDirection = "Output" Then
                        sdata.Adress = "AW" & iOutputStartAdress + (sdata.Kanal * sdata.Kartentyp.OutputAdressLength)
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sdata.Kartentyp.Kartentyp
                    End If
                ElseIf sdata.Kartentyp.Datatyp = "Bit" Then
                    'todo hier fehlt noch die Bit Berechnung
                    If sdata.Kartentyp.SignalDirection = "Input" Then
                        If sdata.Kanal <= 7 Then
                            sdata.Adress = "E" & iInputStartAdress & "." & sdata.Kanal
                        ElseIf sdata.Kanal >= 8 And sdata.Kanal <= 15 Then
                            sdata.Adress = "E" & iInputStartAdress + 1 & "." & sdata.Kanal - 8
                        ElseIf sdata.Kanal >= 16 And sdata.Kanal <= 23 Then
                            sdata.Adress = "E" & iInputStartAdress + 2 & "." & sdata.Kanal - 16
                        ElseIf sdata.Kanal >= 24 And sdata.Kanal <= 31 Then
                            sdata.Adress = "E" & iInputStartAdress + 3 & "." & sdata.Kanal - 24
                        End If
                    ElseIf sdata.Kartentyp.SignalDirection = "Output" Then
                        If sdata.Kanal <= 7 Then
                            sdata.Adress = "A" & iOutputStartAdress & "." & sdata.Kanal
                        ElseIf sdata.Kanal >= 8 And sdata.Kanal <= 15 Then
                            sdata.Adress = "A" & iOutputStartAdress + 1 & "." & sdata.Kanal - 8
                        ElseIf sdata.Kanal >= 16 And sdata.Kanal <= 23 Then
                            sdata.Adress = "A" & iOutputStartAdress + 2 & "." & sdata.Kanal - 16
                        ElseIf sdata.Kanal >= 24 And sdata.Kanal <= 31 Then
                            sdata.Adress = "A" & iOutputStartAdress + 3 & "." & sdata.Kanal - 24
                        End If
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sdata.Kartentyp.Kartentyp
                    End If
                Else
                    MsgBox "Adressen für Kartentyp: " & sdata.Kartentyp.Kartentyp & " kann nicht berechnet werden"
                End If
                ' Ergebnis einsammeln
                rData.Addobj sdata
            Next
            ' Berechnung der Adressen für nächste Karte
            iTmpAdressInput = 0
            iTmpAdressOutput = 0
            x = rData.Count
            If rData.Item(x).Kartentyp.Datatyp = "Byte" Then
                If rData.Item(x).Kartentyp.SignalDirection = "Input" Then
                    iTmpAdressInput = rData.Item(x).Kartentyp.InputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                ElseIf rData.Item(x).Kartentyp.SignalDirection = "Output" Then
                    iTmpAdressOutput = rData.Item(x).Kartentyp.OutputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                ElseIf rData.Item(x).Kartentyp.SignalDirection = "InOut" Then
                    iTmpAdressInput = rData.Item(x).Kartentyp.InputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                    iTmpAdressOutput = rData.Item(x).Kartentyp.OutputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                Else
                    MsgBox "Adressen für Kartentyp: " & rData.Item(x).Kartentyp.Kartentyp & " kann nicht berechnet werden"
                End If
            ElseIf rData.Item(x).Kartentyp.Datatyp = "Word" Then
                If rData.Item(x).Kartentyp.SignalDirection = "Input" Then
                    iTmpAdressInput = rData.Item(x).Kartentyp.InputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                ElseIf rData.Item(x).Kartentyp.SignalDirection = "Output" Then
                    iTmpAdressOutput = rData.Item(x).Kartentyp.OutputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
               ElseIf rData.Item(x).Kartentyp.SignalDirection = "InOut" Then
                    iTmpAdressInput = rData.Item(x).Kartentyp.InputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                    iTmpAdressOutput = rData.Item(x).Kartentyp.OutputAdressLength * (rData.Item(x).Kartentyp.MaxChannelPerSlot + 1)
                Else
                    MsgBox "Adressen für Kartentyp: " & rData.Item(x).Kartentyp.Kartentyp & " kann nicht berechnet werden"
                End If
            ElseIf rData.Item(x).Kartentyp.Datatyp = "Bit" Then
       
                If rData.Item(x).Kartentyp.MaxChannelPerSlot <= 7 Then
                    If rData.Item(x).Kartentyp.SignalDirection = "Input" Then
                        iTmpAdressInput = 1      'Byte
                    ElseIf rData.Item(x).Kartentyp.SignalDirection = "Output" Then
                        iTmpAdressOutput = 1
                    Else
                        MsgBox "Adressen für Kartentyp: " & rData.Item(x).Kartentyp.Kartentyp & " kann nicht berechnet werden"
                    End If
                ElseIf rData.Item(x).Kartentyp.MaxChannelPerSlot > 8 And rData.Item(x).Kartentyp.MaxChannelPerSlot <= 15 Then
                    If rData.Item(x).Kartentyp.SignalDirection = "Input" Then
                        iTmpAdressInput = 2      'Byte
                    ElseIf rData.Item(x).Kartentyp.SignalDirection = "Output" Then
                        iTmpAdressOutput = 2
                    Else
                        MsgBox "Adressen für Kartentyp: " & rData.Item(x).Kartentyp.Kartentyp & " kann nicht berechnet werden"
                    End If
                ElseIf rData.Item(x).Kartentyp.MaxChannelPerSlot > 16 And rData.Item(x).Kartentyp.MaxChannelPerSlot <= 32 Then
                    If rData.Item(x).Kartentyp.SignalDirection = "Input" Then
                        iTmpAdressInput = 4      'Byte
                    ElseIf rData.Item(x).Kartentyp.SignalDirection = "Output" Then
                        iTmpAdressOutput = 4
                    Else
                        MsgBox "Adressen für Kartentyp: " & rData.Item(x).Kartentyp.Kartentyp & " kann nicht berechnet werden"
                    End If
                Else
                    MsgBox "Fehler in Datentyp, Berechnung nicht möglich"
                End If
                            Else
                MsgBox "Fehler in Datentyp, Berechnung nicht möglich"
            End If
            iInputStartAdress = iInputStartAdress + iTmpAdressInput + rData.Item(x).Kartentyp.InputAdressDiagnosticLength
               iOutputStartAdress = iOutputStartAdress + iTmpAdressOutput + rData.Item(x).Kartentyp.OutputAdressDiagnosticLength
        Next y
    End If
    Set AdressPerSlottyp = rData
        
End Function



