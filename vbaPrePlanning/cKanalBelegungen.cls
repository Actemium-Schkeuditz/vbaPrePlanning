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

Public Sub Add(ByRef key As String, ByRef KWSBMK As String, ByRef Signal As Integer, ByRef Stationsnummer As Integer, ByRef Kartentyp As String, Optional ByRef Steckplatz As Integer, Optional ByRef Kanal As Integer, Optional ByRef Segmentvorlage As String, Optional ByRef Adresse As String)
    'create a dataset for Chanel and add to collection

    Dim p As New cBelegung
    p.key = key
    p.KWSBMK = KWSBMK
    p.Signal = Signal
    p.Stationsnummer = Stationsnummer
    p.Kartentyp.Kartentyp = Kartentyp
    p.Steckplatz = Steckplatz
    p.Kanal = Kanal
    p.Segmentvorlage = Segmentvorlage
    p.Adress = Adresse
    
    objKanalBelegungen.Add p
End Sub

Public Sub Addobj(ByRef obj As cBelegung)
    objKanalBelegungen.Add obj
End Sub

Public Sub AddDataSet(ByVal dataSet As cBelegung)
    'create a new dataset and add to collection
    Dim p As New cBelegung
    p.key = dataSet.key
    p.KWSBMK = dataSet.KWSBMK
    p.Signal = dataSet.Signal
    p.Stationsnummer = dataSet.Stationsnummer
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.Steckplatz = dataSet.Steckplatz
    p.Kanal = dataSet.Kanal
    p.Segmentvorlage = dataSet.Segmentvorlage
    p.Adress = dataSet.Adress
    
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


Public Sub ReadExcelDataChanelToCollection(ByRef tablename As String, ByRef dataColection As cKanalBelegungen, ByRef spalteStationsnummer As String, ByRef spalteKartentyp As String)
    'Reads the table on the "EplSheet" sheet and adds
    'Chanels to the collection. It also adds values
    'to the product's properties.
    Dim rCell As Range
    Dim rTable As Range
    Dim iSpalteIntStationsnummer As Integer
    Dim ispalteKartentyp As Integer
    
    On Error GoTo ErrorHandle

    ' Konvertierung Spaltenbuchstaben in INTEGER
    iSpalteIntStationsnummer = SpaltenBuchstaben2Int(spalteStationsnummer) - 1
       
    ispalteKartentyp = SpaltenBuchstaben2Int(spalteKartentyp) - 1
    'First we check if the table is intact. The
    'check is NOT complete.
    ThisWorkbook.Worksheets(tablename).Activate
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
                Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 1, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp).Value, rCell.Offset(0, ispalteKartentyp + 2).Value, rCell.Offset(0, ispalteKartentyp + 3).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4).Value
                'Signal 2
                If rCell.Offset(0, ispalteKartentyp + 12).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, 2, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 12).Value, rCell.Offset(0, ispalteKartentyp + 2 + 12).Value, rCell.Offset(0, ispalteKartentyp + 3 + 12).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 12).Value
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
                If .Cells.Item(i, "A") = sdata.key Then
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
    Dim station As Variant
    
    For Each station In pStationsnummer
        For Each sdata In Me
            If sdata.Stationsnummer = station Then
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

Public Function zuweisenKanal(offsetSlot As Integer, ByVal pKartentyp As String) As cKanalBelegungen

    ' hochzählen der Kanäle und zurückgeben als Class
    Dim iKanal As Integer
    Dim iNextKanal As Integer
    Dim iSteckplatz As Integer

    Dim rData As New cKanalBelegungen
    Dim sortData As New cKanalBelegungen
    Dim sKWSBMKold As String
    Dim bET200SPNothalt As Boolean
    Dim bRunOnce As Boolean
    Dim sSegmentcorrection As String
    
    ' Startwerte
    iKanal = 0
    bET200SPNothalt = False
    bRunOnce = False
    If offsetSlot = 0 Then
        iSteckplatz = 1
    Else
        iSteckplatz = 0
    End If
    
    sKWSBMKold = vbNullString

    Dim sdata As New cBelegung
    If pKartentyp = "ET200SP 8FDI" Then
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sdata In sortData
            ' Kanäle müssen passend aufgelgt werden
            If Left(sdata.Segmentvorlage, 20) = "FDI_ET200SP_NOT_HALT" Or ((sdata.Segmentvorlage = "FDI_ET200SP_ZUHALTUNG_AZM201B") And (sdata.Signal = 1 Or sdata.Signal = 2)) Then
                bET200SPNothalt = True           ' Hier FDI doppelte Auswertung der Signale
                If sdata.Signal = 2 Then
                    ' Kanäle 4, 5, 6 und 7
                    sdata.Kanal = iNextKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    If iKanal <= (sdata.Kartentyp.MaxChannelPerSlot / 2) Then
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + offsetSlot
                        iNextKanal = iKanal + 4
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iNextKanal = iKanal + 4
                        iSteckplatz = iSteckplatz + 1
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + offsetSlot
                    End If
                End If
            ElseIf bET200SPNothalt = True Then
            ' hier weiter wenn einmal ein Not-Aus kam
                If bRunOnce = False Then
                    ' Kanäle 4, 5, 6 und 7
                    sdata.Kanal = iNextKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    bRunOnce = False
                    If iKanal <= (sdata.Kartentyp.MaxChannelPerSlot / 2) Then
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + offsetSlot
                        iNextKanal = iKanal + 4
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iNextKanal = iKanal + 4
                        iSteckplatz = iSteckplatz + 1
                        sdata.Kanal = iKanal
                        sdata.Steckplatz = iSteckplatz + offsetSlot
                    End If
                End If
                'normale Belegung
            Else
                sdata.Kanal = iKanal
                sdata.Steckplatz = iSteckplatz + offsetSlot
                If iKanal < sdata.Kartentyp.MaxChannelPerSlot Then
                    iKanal = iKanal + 1
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                End If
            End If
            sKWSBMKold = sdata.KWSBMK
            rData.AddDataSet sdata
        Next
    ElseIf pKartentyp = "CPX-8DE-D" Then
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sdata In sortData
            'pro Stecker werden zwei Kanäle belegt
            sSegmentcorrection = Right(sdata.Segmentvorlage, 2)
            If iKanal = 1 Or iKanal = 3 Or iKanal = 5 Or iKanal = 7 Then
                'ungerade Kanäle 1, 3, 5 und 7
                If iKanal <= sdata.Kartentyp.MaxChannelPerSlot Then
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
                  
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
                End If
                'sData.Segmentvorlage = "ungerade"
                If sSegmentcorrection = "_1" Then
               sdata.Segmentvorlage = Left(sdata.Segmentvorlage, Len(sdata.Segmentvorlage) - 2) & "_2"
               End If
            Else
                ' Kanäle 0, 2, 4 und 6
                If iKanal < sdata.Kartentyp.MaxChannelPerSlot Then
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
                    iNextKanal = iKanal + 4
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sdata.Kanal = iKanal
                    sdata.Steckplatz = iSteckplatz + offsetSlot
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
        For Each sdata In Me
       
            sdata.Kanal = iKanal
            sdata.Steckplatz = iSteckplatz + offsetSlot
            rData.AddDataSet sdata
            If iKanal < sdata.Kartentyp.MaxChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
            End If
        Next
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

Public Function returnAllSlotsPerRack() As cKanalBelegungen
    ' Rückgabe der Steckplätze mit dem ersten Datensatz, bzw Kanal 0

    ' On Error GoTo ErrorHandler
    Set returnAllSlotsPerRack = Nothing
  '  Dim sData As New cBelegung
 'Dim col As New cKanalBelegungen
   ' Dim col As Variant
    
    
    
    ' Suchen nach allen Datensätzten dem ersten belegten Kanal
    Dim sdata As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sdata In Me
        If sdata.Kanal = 0 Then
            rData.AddDataSet sdata
            'ElseIf sData.Kanal = 1 Then
            'rData.AddDataSet sData
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





