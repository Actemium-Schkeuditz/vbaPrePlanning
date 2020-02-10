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


Private objKanalBelegungen As New Collection
'Private objKanalBelegungen As Collection

'Public Kartentyp As String                'Key
'Public Kanal As Integer               'Secondary key

Private Sub Class_Initialize()
    Set objKanalBelegungen = New Collection
End Sub

Private Sub Class_Terminate()
    Set objKanalBelegungen = Nothing
End Sub

Public Sub Add(KWSBMK As String, Signal As Integer, Stationsnummer As Integer, Kartentyp As String, Optional Steckplatz As Integer, Optional Kanal As Integer, Optional Segmentvorlage As String, Optional Adresse As String)
    'create a dataset for Chanel and add to collection

    Dim p As New cBelegung
    p.KWSBMK = KWSBMK
    p.Signal = Signal
    p.Stationsnummer = Stationsnummer
    p.Kartentyp = Kartentyp
    p.Steckplatz = Steckplatz
    p.Kanal = Kanal
    p.Segmentvorlage = Segmentvorlage
    p.Adress = Adresse
    
    objKanalBelegungen.Add p
End Sub


Public Sub Addobj(obj As cBelegung)
    objKanalBelegungen.Add obj
End Sub

Public Sub AddDataSet(dataSet As cBelegung)
    'create a new dataset and add to collection
    Dim p As New cBelegung
    p.KWSBMK = dataSet.KWSBMK
    p.Signal = dataSet.Signal
    p.Stationsnummer = dataSet.Stationsnummer
    p.Kartentyp = dataSet.Kartentyp
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

Public Sub Remove(NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objKanalBelegungen.Remove NameOrNumber
End Sub
'Attribute Item.VB_UserMemId = 0
Public Property Get Item(Index As Variant) As cBelegung
Attribute Item.VB_UserMemId = 0
    Set Item = objKanalBelegungen.Item(Index)
End Property

'VBA BUG falls die Suchfunktion nicht funktioniert folgendes hier einfügen:
'Attribute NewEnum.VB_UserMemId = -4
'danach im Notepad die exortierte Klasse bearbeiten, die Auskommentierung für: "' Attribute NewEnum.VB_UserMemId = -4" entfernen und wieder importieren
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objKanalBelegungen.[_NewEnum]
End Property

Public Function Sort() As cKanalBelegungen
    Dim i As Long, j As Long, k As Long, bln As Boolean
    Dim lngCount As Long, arr() As Long, ppl As cKanalBelegungen
    lngCount = Me.Count
    If lngCount > 0 Then
        ReDim arr(0 To lngCount - 1)
        For i = 0 To lngCount - 1: arr(i) = i + 1: Next

        For i = 1 To lngCount - 1
            k = arr(i)
            j = i - 1
            bln = False
            Do
                If Me(arr(j)).CompareTo(Me(k)) > 0 Then
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
    For i = 0 To lngCount - 1: ppl.Addobj Me(arr(i)): Next

    Set Sort = ppl
End Function


Public Sub ReadExcelDataChanelToCollection(tablename As String, dataColection As cKanalBelegungen, spalteStationsnummer As String, spalteKartentyp As String)
    'Reads the table on the "EplSheet" sheet and adds
    'Chanels to the collection. It also adds values
    'to the product's properties.
    Dim rCell As Range
    Dim rTable As Range
    Dim sEinbauort As String
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
        'Set EinbauorteData = Nothing
        Set dataColection = Nothing
        Exit Sub
    End If

    Set rTable = Range("A3")

    With rTable
        If Len(.Offset(0, 0).Value) = 0 Or Len(.Offset(1, 0).Value) = 0 Then
            MsgBox "The products table has been destoyed and there is no or too few data.", vbCritical, "Error"
            ' Set EinbauorteData = Nothing
            Set dataColection = Nothing
            Exit Sub
        End If
    End With

    'If the table is okay, we start reading data.
    'The range is expanded to include the last
    'cell with content.
    If Len(rTable.Offset(2, 0).Formula) = 0 Then
        Set rTable = rTable.Offset(1, 9)
    Else
        Set rTable = Range(rTable.Offset(1, 0), rTable.End(xlDown))
    End If

    For Each rCell In rTable
        With rCell
            ' prüfen ob Stationsnummer nicht leer
            If rCell.Offset(0, iSpalteIntStationsnummer).Value <> vbNullString Then
                '      sEinbauort = str$(.Value)
                ' die gelesenen Daten anhängen
                'Me.Add CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp).Value, rCell.Offset(0, ispalteKartentyp + 2).Value, rCell.Offset(0, ispalteKartentyp + 3).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4).Value
                Me.Add rCell.Offset(0, 1).Value, 1, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp).Value, rCell.Offset(0, ispalteKartentyp + 2).Value, rCell.Offset(0, ispalteKartentyp + 3).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4).Value
                'Signal 2
                If rCell.Offset(0, ispalteKartentyp + 12).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 1).Value, 2, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 12).Value, rCell.Offset(0, ispalteKartentyp + 2 + 12).Value, rCell.Offset(0, ispalteKartentyp + 3 + 12).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 12).Value
                End If
                'Signal 3
                If rCell.Offset(0, ispalteKartentyp + 24).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 1).Value, 3, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 24).Value, rCell.Offset(0, ispalteKartentyp + 2 + 24).Value, rCell.Offset(0, ispalteKartentyp + 3 + 24).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 24).Value
                End If
                'Signal 4
                If rCell.Offset(0, ispalteKartentyp + 36).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 1).Value, 4, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 36).Value, rCell.Offset(0, ispalteKartentyp + 2 + 36).Value, rCell.Offset(0, ispalteKartentyp + 3 + 36).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 36).Value
                End If
                'Signal 5
                If rCell.Offset(0, ispalteKartentyp + 48).Value <> vbNullString Then
                    Me.Add rCell.Offset(0, 1).Value, 5, CInt(rCell.Offset(0, iSpalteIntStationsnummer).Value), rCell.Offset(0, ispalteKartentyp + 48).Value, rCell.Offset(0, ispalteKartentyp + 2 + 48).Value, rCell.Offset(0, ispalteKartentyp + 3 + 48).Value, rCell.Offset(0, ispalteKartentyp - 13).Value, rCell.Offset(0, ispalteKartentyp + 4 + 48).Value
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
    MsgBox Err.Description & " Procedure MakeProductCollection.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Function searchKanalBelegungenKartentyp(ByRef pStationsnummer As Collection, pKartentyp As String, data As cKanalBelegungen) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Kartentyp
    Dim sData As New cBelegung
    'searchAnschluss = Nothing
    Dim station As Variant
    
    For Each station In pStationsnummer
        For Each sData In data
            If sData.Stationsnummer = station Then
                If sData.Kartentyp = pKartentyp Then
                    Me.AddDataSet sData
                    'Debug.Print searchAnschluss.Item(1).Kartentyp;
                    '    Exit For                         ' nur einmal suchen dann beenden
                End If
            End If
        Next
    Next
    If Me.Count = 0 Then
        Me.Add vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    'Test der Suche
    '  Debug.Print searchAnschluss.Item(1).Kartentyp; vbTab; searchAnschluss.Item(1).Kanal; vbTab; searchAnschluss.Item(1).Anschluss_1; vbTab; searchAnschluss.Item(1).Anschluss_2; vbTab; searchAnschluss.Item(1).Anschluss_3; vbTab; searchAnschluss.Item(1).Anschluss_4; vbTab; searchAnschluss.Item(1).Anschluss_M; vbTab; searchAnschluss.Item(1).Anschluss_VS
End Function


Public Function zuweisenKanal(data As cKanalBelegungen) As cKanalBelegungen
    ' hochzählen der Kanäle und zurückgeben as Class
    Dim iKanal As Integer
    Dim iStationOld As Integer
    'Dim rData As cKanalBelegungen
    iKanal = 0
    iStationOld = 0
    Dim sData As New cBelegung
    For Each sData In data
    If iStationOld = 0 Or iStationOld = sData.Stationsnummer Then
        iKanal = iKanal + 1
        Else
        iKanal = 0
        End If
        sData.Kanal = iKanal
        sData.Steckplatz = 0
        Me.AddDataSet sData
        '       rData.AddDataSet sData
        
        iStationOld = sData.Stationsnummer
    Next
    'Set zuweisenKanal = rData
End Function




