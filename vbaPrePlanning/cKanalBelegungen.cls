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

Public Sub Add(ByRef Key As String, ByRef KWSBMK As String, ByVal Signal As Integer, ByRef Stationsnummer As Integer, Optional ByRef Kartentyp As String, _
               Optional ByVal Steckplatz As Integer, Optional ByRef Kanal As Integer, Optional ByRef Segmentvorlage As String, Optional ByRef Adresse As String, Optional ByRef InputAdressSize As Long, _
               Optional ByRef OutputAdressSize As Long, Optional ByRef SPSBMK As String, Optional ByRef Anlage As String, Optional ByRef Einbauort As String, Optional ByRef SubSlot As Integer, _
               Optional ByRef EinbauortEinzel As String)
    'create a dataset for Chanel and add to collection

    Dim p As New cBelegung
    p.Key = Key
    p.KWSBMK = KWSBMK
    p.Signal = Signal
    p.Stationsnummer = Stationsnummer
    p.Kartentyp.Kartentyp = Kartentyp
    p.Steckplatz = Steckplatz
    p.SubSlot = SubSlot
    p.Kanal = Kanal
    p.Segmentvorlage = Segmentvorlage
    p.Adress = Adresse
    p.InputAdressSize = InputAdressSize
    p.OutputAdressSize = OutputAdressSize
    p.SPSBMK = SPSBMK
    p.Anlage = Anlage
    p.Einbauort = Einbauort
    p.EinbauortEinzel = EinbauortEinzel
    
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
    p.Anlage = dataSet.Anlage
    p.Einbauort = dataSet.Einbauort
    p.EinbauortEinzel = dataSet.EinbauortEinzel
    p.Seite = dataSet.Seite
    p.Signal = dataSet.Signal
    p.Stationsnummer = dataSet.Stationsnummer
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.Steckplatz = dataSet.Steckplatz
    p.SubSlot = dataSet.SubSlot
    p.Kanal = dataSet.Kanal
    p.Segmentvorlage = dataSet.Segmentvorlage
    p.Adress = dataSet.Adress
    p.InputAdressSize = dataSet.InputAdressSize
    p.OutputAdressSize = dataSet.OutputAdressSize
    p.InputAdressPerSlot = dataSet.InputAdressPerSlot
    p.OutputAdressPerSlot = dataSet.OutputAdressPerSlot
    p.Anschluss1 = dataSet.Anschluss1
    p.Anschluss2 = dataSet.Anschluss2
    p.Anschluss3 = dataSet.Anschluss3
    p.Anschluss4 = dataSet.Anschluss4
    p.AnschlussM = dataSet.AnschlussM
    p.AnschlussVS = dataSet.AnschlussVS
    p.OutputAdress = dataSet.OutputAdress
    
    p.SPSBMK = dataSet.SPSBMK
    p.SymbolischeAdresse = dataSet.SymbolischeAdresse
    
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

Public Function SortForPageNumbers() As cKanalBelegungen
    ' sort higher level function, mounting location and KWSBMK
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
                If Me.Item(arr(j)).CompareForPagenumbers(Me.Item(k)) > 0 Then
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

    Set SortForPageNumbers = ppl
End Function

Public Sub ReadExcelDataChanelToCollection(ByRef tablename As String, ByRef dataColection As cKanalBelegungen)
    'Reads the table on the "EplSheet" sheet and adds
    'Chanels to the collection. It also adds values
    'to the product's properties.
    Dim rCell As Range
    Dim rTable As Range
    Dim iSpalteEingangsAdress As Long
    Dim iSpalteAusgangsAdress As Long
    Dim iSpalteStationsnummer As Long
    Dim iSpalteSymbolischeAdresse As Long
    Dim iSpalteKartentyp As Long
    Dim iSpalteSPSBMK As Long
    Dim iSpalteSegmentvorlage As Long
    Dim iSpalteSegmentdefinition As Long
    Dim iSpalteEinbauortEinzel As Long
    Dim iSpalteEinbauort As Long
    Dim iSpalteSteckplatz As Long
    Dim rData As New cBelegung
    Dim iSignal As Long
    
    Dim ExcelConfig As New cExcelConfig
    
    On Error GoTo ErrorHandle

    iSpalteStationsnummer = SpaltenBuchstaben2Int(ExcelConfig.StationsnummerSignal_1) - 1
    iSpalteEingangsAdress = SpaltenBuchstaben2Int(ExcelConfig.EingangsAdress) - 1
    iSpalteAusgangsAdress = SpaltenBuchstaben2Int(ExcelConfig.AusgangsAdress) - 1
    iSpalteSymbolischeAdresse = SpaltenBuchstaben2Int(ExcelConfig.SympolischeAdresse) - 1
    iSpalteSegmentvorlage = SpaltenBuchstaben2Int(ExcelConfig.Segmentvorlage) - 1
    iSpalteEinbauortEinzel = SpaltenBuchstaben2Int(ExcelConfig.EinbauortEinzel) - 1
    iSpalteEinbauort = SpaltenBuchstaben2Int(ExcelConfig.Einbauort) - 1
    iSpalteKartentyp = SpaltenBuchstaben2Int(ExcelConfig.Kartentyp) - 1
    iSpalteSPSBMK = SpaltenBuchstaben2Int(ExcelConfig.SPSBMK) - 1
    iSpalteSteckplatz = SpaltenBuchstaben2Int(ExcelConfig.Steckplatz) - 1
    iSpalteSegmentdefinition = SpaltenBuchstaben2Int(ExcelConfig.Segmentdefinition) - 1
    
    
    iSignal = 1
    'First we check if the table is intact. The
    'check is NOT complete.
    ThisWorkbook.Worksheets.[_Default](tablename).Activate
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
            ' prüfen ob Stationsnummer nicht leer, oder Datensätze vom Typ "Schaltschrankgeräte" oder "SPSundNetzwerk"
            If rCell.Offset(0, iSpalteStationsnummer).Value2 <> vbNullString Or rCell.Offset(0, iSpalteSegmentdefinition).Value2 = "Schaltschrankgeräte" Or rCell.Offset(0, iSpalteSegmentdefinition).Value2 = "SPSundNetzwerk" Then
                ' die gelesenen Daten anhängen
                For iSignal = 1 To 6
                    If rCell.Offset(0, iSpalteKartentyp - 1 + (14 * (iSignal - 1))).Value <> vbNullString Then
                        rData.Key = rCell.Offset(0, 0).Value
                        rData.KWSBMK = rCell.Offset(0, 1).Value2
                        rData.Anlage = rCell.Offset(0, 2).Value2
                        rData.SymbolischeAdresse = rCell.Offset(0, iSpalteSymbolischeAdresse).Value2
                        rData.Einbauort = rCell.Offset(0, iSpalteEinbauort).Value
                        rData.EinbauortEinzel = rCell.Offset(0, iSpalteEinbauortEinzel).Value
                        rData.Signal = iSignal
                        rData.Stationsnummer = CInt(rCell.Offset(0, iSpalteStationsnummer + (14 * (iSignal - 1))).Value2)
                    
                        rData.Kartentyp.Kartentyp = rCell.Offset(0, iSpalteKartentyp + (14 * (iSignal - 1))).Value
                        rData.Steckplatz = CInt(rCell.Offset(0, iSpalteSteckplatz + (14 * (iSignal - 1))).Value2)
                        rData.Kanal = CInt(rCell.Offset(0, iSpalteKartentyp + 3 + (14 * (iSignal - 1))).Value)
                    
                    
                        rData.Segmentvorlage = rCell.Offset(0, iSpalteSegmentvorlage).Value
                        rData.Adress = rCell.Offset(0, iSpalteKartentyp + 4 + (14 * (iSignal - 1))).Value
                        rData.InputAdressSize = CLng(rCell.Offset(0, iSpalteEingangsAdress))
                        rData.OutputAdressSize = CLng(rCell.Offset(0, iSpalteAusgangsAdress))
                    
                        rData.SPSBMK = rCell.Offset(0, iSpalteSPSBMK + (14 * (iSignal - 1))).Value
                    
                        Me.AddDataSet rData
                    End If
                Next iSignal
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
    Dim sData As New cBelegung
    Dim ws1 As Worksheet
    Dim iSpalteSteckplatz As Long
    Dim iSpalteKanal As Long
    Dim i As Long
    Dim zeilenanzahl As Long
    Dim ExcelConfig As New cExcelConfig
        
     On Error GoTo ErrorHandle
     
    ' Spalten definieren
    ' nummerischen Wert ermitteln
    iSpalteSteckplatz = SpaltenBuchstaben2Int(ExcelConfig.Steckplatz)
    iSpalteKanal = SpaltenBuchstaben2Int(ExcelConfig.Kanal)
        
    ThisWorkbook.Worksheets.[_Default](tablename).Activate
    Set ws1 = ThisWorkbook.ActiveSheet
    ' Tabelle mit Daten bearbeiten
    With ws1
        
        ' Worksheet activate
        ws1.Activate

        Application.ScreenUpdating = False
        '##### Herausfinden der Anzahl der Zeilen #####
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        
        For Each sData In Me
            For i = 3 To zeilenanzahl
                ' suchen nach dem passenden key
                If .Cells.Item(i, "A") = sData.Key Then
                    .Cells.Item(i, ExcelConfig.SympolischeAdresse) = sData.SymbolischeAdresse
                    .Cells.Item(i, (iSpalteSteckplatz - 1 + (14 * (sData.Signal - 1)))) = sData.SPSBMK
                    .Cells.Item(i, (iSpalteSteckplatz + (14 * (sData.Signal - 1)))) = sData.Steckplatz
                    .Cells.Item(i, (iSpalteKanal + (14 * (sData.Signal - 1)))) = sData.Kanal
                    .Cells.Item(i, (iSpalteKanal + 1 + (14 * (sData.Signal - 1)))) = sData.Adress
                    .Cells.Item(i, (iSpalteKanal + 2 + (14 * (sData.Signal - 1)))) = sData.Anschluss1
                    .Cells.Item(i, (iSpalteKanal + 3 + (14 * (sData.Signal - 1)))) = sData.Anschluss2
                    .Cells.Item(i, (iSpalteKanal + 4 + (14 * (sData.Signal - 1)))) = sData.Anschluss3
                    .Cells.Item(i, (iSpalteKanal + 5 + (14 * (sData.Signal - 1)))) = sData.Anschluss4
                    .Cells.Item(i, (iSpalteKanal + 6 + (14 * (sData.Signal - 1)))) = sData.AnschlussM
                    .Cells.Item(i, (iSpalteKanal + 7 + (14 * (sData.Signal - 1)))) = sData.AnschlussVS
                    .Cells.Item(i, ExcelConfig.Segmentvorlage) = sData.Segmentvorlage
                End If
            
            Next i
        Next
    End With

BeforeExit:
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim schreiben der Datensätze.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Sub writePageNumbersToExcel(ByRef tablename As String)
    Dim sData As New cBelegung
    Dim ws1 As Worksheet

    Dim i As Long
    Dim zeilenanzahl As Long
    Dim ExcelConfig As New cExcelConfig
    On Error GoTo ErrorHandle
        
    ThisWorkbook.Worksheets.[_Default](tablename).Activate
    Set ws1 = ThisWorkbook.ActiveSheet
    ' Tabelle mit Daten bearbeiten
    With ws1
        
        ' Worksheet activate
        ws1.Activate

        Application.ScreenUpdating = False
        '##### Herausfinden der Anzahl der Zeilen #####
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        
        'alte Daten löschen
        For i = 3 To zeilenanzahl
            .Cells.Item(i, ExcelConfig.Seite) = vbNullString
        Next i
        
        For Each sData In Me
            For i = 3 To zeilenanzahl
                ' suchen nach dem passenden key
                If .Cells.Item(i, "A") = sData.Key Then
                    .Cells.Item(i, ExcelConfig.Seite) = sData.Seite
                End If
            Next i
        Next
    End With

BeforeExit:
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim einlesen.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Function searchKanalBelegungenKartentyp(ByRef pStationsnummer As Collection, ByVal pKartentyp As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Kartentyp und sortiert nach Stationsnummer
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
    Set searchKanalBelegungenKartentyp = Nothing
    Dim Station As Variant
    
    For Each Station In pStationsnummer
        For Each sData In Me
            If sData.Stationsnummer = Station Then
                If sData.Kartentyp.Kartentyp = pKartentyp Then
                    rData.AddDataSet sData
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
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
    Set searchKanalBelegungenSteckplatz = Nothing
    'Dim Station As Variant
    
    'For Each Station In pStationsnummer
    For Each sData In Me
        If sData.Stationsnummer = pStationsnummer Then
            If sData.Steckplatz = pSteckplatz Then
                rData.AddDataSet sData
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
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sData In Me
        If sData.Stationsnummer = pStationsnummer Then
            rData.AddDataSet sData
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetPerStation = rData

End Function

Public Function searchDatasetPlcModules(ByVal pPlcModules As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Kartentyp
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sData In Me
        If sData.Kartentyp.Kartentyp = pPlcModules Then
            rData.AddDataSet sData
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetPlcModules = rData
   
End Function

Public Function searchDatasetPlcTyp(ByVal pPLCtyp As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Stationstyp
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sData In Me
        If sData.Kartentyp.PLCtyp = pPLCtyp Then
            rData.AddDataSet sData
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetPlcTyp = rData
   
End Function

Public Function excludeDatasetPlcTyp(ByVal pPLCtyp As String) As cKanalBelegungen
    ' Suchen nach allen Datensätzten die nicht den gesuchten Stationstyp enthalten
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sData In Me
        If sData.Kartentyp.PLCtyp <> pPLCtyp Then
            rData.AddDataSet sData
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set excludeDatasetPlcTyp = rData
   
End Function
Public Function searchDatasetSignal(ByVal pSignal As Integer) As cKanalBelegungen
    ' Suchen nach allen Datensätzten mit einem bestimmten Stationstyp
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
   
    For Each sData In Me
        If sData.Signal = pSignal Then
            rData.AddDataSet sData
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchDatasetSignal = rData
   
End Function

Public Function zuweisenKanal(ByRef OffsetSlot As Integer, ByVal pKartentyp As String, ByVal dataPLCConfig As cPLCconfig) As cKanalBelegungen

    ' hochzählen der Kanäle und zurückgeben als Class
    Dim iKanal As Integer
    Dim iOldKanal As Integer
    Dim iSteckplatz As Integer
    Dim rData As New cKanalBelegungen
    Dim sortData As New cKanalBelegungen
    Dim bET200SPNothalt As Boolean
    Dim bRunOnce As Boolean
    Dim sSegmentCorrection As String
    Dim iReserveChannelPerSlot As Long
    Dim iReserveChannelsBefor As Long
    Dim iReserveChannelsAfter As Long
    Dim iReserveAfterPLCTyp As Long
    Dim i As Long
    
    i = 1
    iReserveChannelPerSlot = dataPLCConfig.Item(i).ReserveChannelPerSlot
    iReserveChannelsBefor = dataPLCConfig.Item(i).ReserveChannelsBefor
    iReserveChannelsAfter = dataPLCConfig.Item(i).ReserveChannelsAfter
    iReserveAfterPLCTyp = dataPLCConfig.Item(i).ReserveSlot
    
    ' Startwerte
    iKanal = 0
    bET200SPNothalt = False
    bRunOnce = False
    iOldKanal = 0
    
    'Kopfstation dazu
    If (pKartentyp = "CPX-8DE-D" Or pKartentyp = "CPX-4DA" Or pKartentyp = "CPX-4AE-T" Or pKartentyp = "CPX 2x3/2 mono" Or pKartentyp = "CPX 5/2 mono" Or pKartentyp = "CPX 5/2 bistabil") And (iSteckplatz + OffsetSlot) = 0 Then
        rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, "FB43", iSteckplatz + OffsetSlot, iKanal
    End If
        
    If OffsetSlot = 0 Then
        iSteckplatz = 1
    Else
        iSteckplatz = 0
    End If
    

    Dim sData As New cBelegung
    If pKartentyp = "ET200SP 8FDI" Then
        'erstmal die Kanäle davor bearbeiten
        Do While iReserveChannelsBefor > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                If iKanal = 0 Then               'And rdata Is Nothing Then
                    rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
                End If
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
            End If
            iReserveChannelsBefor = iReserveChannelsBefor - 1
        Loop
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sData In sortData
            ' Kanäle müssen passend aufgelgt werden
            If Left(sData.Segmentvorlage, 20) = "FDI_ET200SP_NOT_HALT" Or sData.Segmentvorlage = "FDI_ET200SP_BMA" _
                                                Or ((Left(sData.Segmentvorlage, 21) = "FDI_ET200SP_ZUHALTUNG") And (sData.Signal = 1 Or sData.Signal = 2)) _
                                                Or ((Left(sData.Segmentvorlage, 36) = "FDI_ET200SP_SICHERHEITSSCHALTER_2FDI") And (sData.Signal = 1 Or sData.Signal = 2)) Then
                bET200SPNothalt = True           ' Hier FDI doppelte Auswertung der Signale
                If sData.Signal = 2 Then
                    ' Kanäle 4, 5, 6 und 7
                    sData.Kanal = iOldKanal + 4
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    If iKanal <= ((sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) / 2) Then
                        sData.Kanal = iKanal
                        sData.Steckplatz = iSteckplatz + OffsetSlot
                        iOldKanal = iKanal
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iOldKanal = iKanal
                        iSteckplatz = iSteckplatz + 1
                        sData.Kanal = iKanal
                        sData.Steckplatz = iSteckplatz + OffsetSlot
                        iKanal = iKanal + 1
                    End If
                End If
            ElseIf bET200SPNothalt = True Then
                ' hier weiter wenn einmal ein Not-Aus kam
                If bRunOnce = False Then
                    ' Kanäle 4, 5, 6 und 7
                    sData.Kanal = iOldKanal + 4
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                    bRunOnce = True
                Else
                    ' Kanäle 0, 1, 2 und 3
                    bRunOnce = False
                    If iKanal <= ((sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) / 2) Then
                        sData.Kanal = iKanal
                        sData.Steckplatz = iSteckplatz + OffsetSlot
                        iOldKanal = iKanal
                        iKanal = iKanal + 1
                    Else
                        iKanal = 0
                        iOldKanal = iKanal
                        iSteckplatz = iSteckplatz + 1
                        sData.Kanal = iKanal
                        sData.Steckplatz = iSteckplatz + OffsetSlot
                        iKanal = iKanal + 1
                    End If
                End If
                'normale Belegung
            Else
                sData.Kanal = iKanal
                sData.Steckplatz = iSteckplatz + OffsetSlot
                If iKanal < sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                    iKanal = iKanal + 1
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                End If
            End If
            rData.AddDataSet sData
        Next
        'die  nachfolgenden Kanäle bearbeiten
        Do While iReserveChannelsAfter > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
            End If
            iReserveChannelsAfter = iReserveChannelsAfter - 1
            rData.Add "RESERVE", "RESERVE", "1", dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
        Loop
        
    ElseIf pKartentyp = "CPX-8DE-D" Then
        
        Do While iReserveChannelsBefor > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                If iKanal = 0 Then               'And rdata Is Nothing Then
                    rData.Add "LEER1", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
                End If
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
            End If
            iReserveChannelsBefor = iReserveChannelsBefor - 1
        Loop
        
        'Sortierung nach Segmentvorlage
        Set sortData = Me.SortSegment
        For Each sData In sortData
            'pro Stecker werden zwei Kanäle belegt
            sSegmentCorrection = Right(sData.Segmentvorlage, 2)
            If iKanal = 1 Or iKanal = 3 Or iKanal = 5 Or iKanal = 7 Then
                'ungerade Kanäle 1, 3, 5 und 7
                If iKanal <= (sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) Then
                    sData.Kanal = iKanal
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sData.Kanal = iKanal
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                End If
                If sSegmentCorrection = "_1" Then
                    sData.Segmentvorlage = Left(sData.Segmentvorlage, Len(sData.Segmentvorlage) - 2) & "_2"
                End If
            Else
                ' Kanäle 0, 2, 4 und 6
                If iKanal < (sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot) Then
                    sData.Kanal = iKanal
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                    iOldKanal = iKanal + 4
                Else
                    iKanal = 0
                    iSteckplatz = iSteckplatz + 1
                    sData.Kanal = iKanal
                    sData.Steckplatz = iSteckplatz + OffsetSlot
                End If
                If sSegmentCorrection = "_2" Then
                    sData.Segmentvorlage = Left(sData.Segmentvorlage, Len(sData.Segmentvorlage) - 2) & "_1"
                End If
       
            End If
            rData.AddDataSet sData
            iKanal = iKanal + 1
        Next
        'die  nachfolgenden Kanäle bearbeiten
        Do While iReserveChannelsAfter > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
            End If
            iReserveChannelsAfter = iReserveChannelsAfter - 1
        Loop
              
    ElseIf pKartentyp = "ET200AL 16DI" Then
        
        'nicht belegen, sondern nur durchreichen
        For Each sData In Me
            rData.AddDataSet sData
        Next
        
    Else
        'hier bearbeiten alles was keine extra Behandlung benötigt
        'erstmal die Kanäle davor bearbeiten
        Do While iReserveChannelsBefor > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                If iKanal = 0 Then               'And rdata Is Nothing Then
                    rData.Add "LEER1", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
                End If
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                If dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot = 0 And dataPLCConfig.Item(i).Kartentyp.PLCtyp <> "FESTO MPA" Then
                    rData.Add "LEER1", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
                End If
            End If
            iReserveChannelsBefor = iReserveChannelsBefor - 1
        Loop
        
        For Each sData In Me
            sData.Kanal = iKanal
            If sData.Kartentyp.PLCtyp = "IFM IO-LINK MASTER" Then
                If iKanal = 0 Then
                    rData.Add "DIAG_IFM", "DIAGNOSE", 0, sData.Stationsnummer, "AL140x_DIAG", 0, 0
                End If
                sData.SubSlot = iKanal + 1
            ElseIf pKartentyp = "TBEN-S-RFID" Then
                rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, "TBEN-S-HF-Busmodul", iSteckplatz + OffsetSlot, iKanal
            End If
            
            sData.Steckplatz = iSteckplatz + OffsetSlot
            rData.AddDataSet sData
            If iKanal < sData.Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
            End If
        Next
        'die  nachfolgenden Kanäle bearbeiten
        Do While iReserveChannelsAfter > 0
            If iKanal < dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot - iReserveChannelPerSlot Then
                If iKanal = 0 Then
                    rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
                End If
                iKanal = iKanal + 1
            Else
                iKanal = 0
                iSteckplatz = iSteckplatz + 1
                rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, dataPLCConfig.Item(i).Kartentyp.Kartentyp, iSteckplatz + OffsetSlot, iKanal
            End If
            iReserveChannelsAfter = iReserveChannelsAfter - 1
        Loop
    End If
    ' hier die Reserven und Diagnosen anfügen
    If pKartentyp = "TBEN-S-RFID" Then
        rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, "TBEN-S-DIAG", iSteckplatz + OffsetSlot, iKanal
    End If
    
    If dataPLCConfig.Item(i).Kartentyp.PLCtyp = "IFM IO-LINK MASTER" Then
        Do While dataPLCConfig.Item(i).Kartentyp.MaxChannelPerSlot > rData.Count - 2
            rData.Add "IFM_RESERVE", "IFM_RESERVE", 0, dataPLCConfig.Item(i).Stationsnummer, "AL140x_RES", iSteckplatz, iKanal, vbNullString, vbNullString, 0, 0, vbNullString, vbNullString, vbNullString, iKanal + 1
            iKanal = iKanal + 1
        Loop
    End If
    
    Do While iReserveAfterPLCTyp >= 1
        If iKanal > 0 Then
            iSteckplatz = iSteckplatz + 1
        End If
        If dataPLCConfig.Item(i).Kartentyp.PLCtyp = "ET200SP" Then
            rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, "ET200SP RESERVE", iSteckplatz + OffsetSlot, 0
        ElseIf dataPLCConfig.Item(i).Kartentyp.PLCtyp = "FESTO MPA" Then
            rData.Add "LEER", "LEER", 0, dataPLCConfig.Item(i).Stationsnummer, "CPX RESERVE", iSteckplatz + OffsetSlot, 0
        End If
        iReserveAfterPLCTyp = iReserveAfterPLCTyp - 1
        iKanal = 1
    Loop
  
    Set zuweisenKanal = rData
End Function

Public Function returnStation() As Collection
    ' collect all Stations without duplicates
    On Error GoTo ErrorHandler
    Set returnStation = Nothing

    Dim col As New Collection
    Dim bSearchinCol As Boolean
    Dim it As Variant
    Dim sData As New cBelegung
    Dim data As New cKanalBelegungen
    
    Set data = Me.SortChannel
    
    For Each sData In data
        ' prüfen ob Stationsnummer schon in Collection
        bSearchinCol = True
        For Each it In col
            If it = sData.Stationsnummer Then
                bSearchinCol = False
                
            End If
        Next
        If bSearchinCol = True Then
            If sData.Stationsnummer <> 0 Then
                col.Add sData.Stationsnummer     '  dynamically add value to the end
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
    Dim sData As New cBelegung
    
    For Each sData In Me
        ' prüfen ob Kartentyp schon in Collection
        bSearchinCol = True
        For Each it In col
            If it = sData.Kartentyp.Kartentyp Then
                bSearchinCol = False
                
            End If
        Next
        If bSearchinCol = True Then
            If sData.Kartentyp.Kartentyp <> vbNullString Then
                col.Add sData.Kartentyp.Kartentyp          '  dynamically add value to the end
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
    Dim sData As New cBelegung
    maxSlotNumber = 0

    For Each sData In Me
        If sData.Steckplatz > maxSlotNumber Then
        maxSlotNumber = sData.Steckplatz
        End If
    Next
    returnLastSlotNumber = maxSlotNumber
End Function

Public Function returnFirstSlotNumber() As Integer

    Dim minSlotNumber As Integer
    Dim sData As New cBelegung
   
    minSlotNumber = 1000

    For Each sData In Me
        If minSlotNumber > sData.Steckplatz Then
        minSlotNumber = sData.Steckplatz
        End If
    
    Next
    returnFirstSlotNumber = minSlotNumber
End Function

Public Function returnAllSlotsPerRack() As cKanalBelegungen
    ' Rückgabe der Steckplätze mit dem ersten Datensatz, bzw Kanal 0
    
    On Error GoTo ErrorHandler
    Set returnAllSlotsPerRack = Nothing
    Dim icount As Long
   
    ' Suchen nach allen Datensätzten dem ersten belegten Kanal
    Dim sData As New cBelegung
    Dim rData As New cKanalBelegungen
    Set rData = Nothing
   
    For Each sData In Me
        If rData Is Nothing Then
            icount = 0
        Else
            icount = rData.Count
        End If
    
        If sData.Kanal = 0 Then
            rData.AddDataSet sData
            'Exit For
        ElseIf sData.Kanal = 1 And Not (rData Is Nothing) And icount >= 1 Then
            If rData.Item(icount).Steckplatz <> sData.Steckplatz Then
                rData.AddDataSet sData
            End If
        Else
            ' MsgBox " Karte ohne Belegung, Stationsnummer" & sData.Stationsnummer & "Steckplatz: " & sData.Steckplatz
        End If
    Next
 
    If Me.Count = 0 Then
        rData.Add vbNullString, vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set returnAllSlotsPerRack = rData

    Exit Function
ErrorHandler:
    If Err.Number > 0 Then
        MsgBox "Steckplätze konnten nicht ermittelt werden "
        Err.Clear
        Resume Next
    End If
End Function


Public Function AdressPerSlottyp(ByRef iInputStartAdress As Long, ByRef iOutputStartAdress As Long, ByVal iStationsnummer As Integer, ByVal pKartentyp As String) As cKanalBelegungen
    'Berechnen der Adressen, wegen der Variabilität der IO-Link Geräte nur möglich wenn die gesamte Station bzw. der Kartentyp belegt ist
    
    Dim iTmpAdressInput As Long
    Dim iTmpAdressOutput As Long
    Dim iAdressVariInput As Long
    Dim iAdressVariOutput As Long
    Dim ifirstSlot As Integer
    Dim ilastSlot As Integer
    Dim iSubSlot As Long
    Dim sData As New cBelegung
    Dim sDataFesto As New cBelegung
    Dim sDataFestoSearch As New cBelegung
    Dim sDataTurck As New cBelegung
    Dim sdataVari As New cBelegung
    Dim SlotData As New cKanalBelegungen
    Dim rData As New cKanalBelegungen
    Dim rDataSearch As New cKanalBelegungen
    Dim i As Long
    Dim y As Long
    Dim x As Long
    Dim iAdressLegthInput As Long
    Dim iAdressLegthOutput As Long
    
    
    iAdressLegthInput = 0
    iAdressLegthOutput = 0

    ' sortieren nach Station, Slot and Channel
    Me.SortChannel
   
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
                If SlotData.Item(1).Kartentyp.Kartentyp = "ET200SP 4IO-LINK" Then
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
                        MsgBox "Fehler in der Adresslänge von: " & sData.KWSBMK
                    End If
                End If
                iAdressVariInput = iInputStartAdress
                iAdressVariOutput = iOutputStartAdress
            ElseIf pKartentyp = "AL1400" Or pKartentyp = "AL1402" Then 'Berechnung pro Kanal und dann zusammen zählen
                iSubSlot = 0
                iAdressVariInput = iInputStartAdress
                iAdressVariOutput = iOutputStartAdress
                For Each sdataVari In SlotData
                    sdataVari.InputAdressPerSlot = iInputStartAdress
                    sdataVari.OutputAdressPerSlot = iOutputStartAdress
                    
                    sdataVari.Adress = "EB" & iAdressVariInput
                    sdataVari.OutputAdress = "AB" & iAdressVariOutput
                    If sdataVari.Kartentyp.Kartentyp = "AL140x_DIAG" Or sdataVari.Kartentyp.Kartentyp = "AL140x_RES" Then
                        iAdressVariInput = iAdressVariInput + sdataVari.Kartentyp.InputAdressLength + sdataVari.Kartentyp.InputAdressDiagnosticLength
                        iAdressVariOutput = iAdressVariOutput + sdataVari.Kartentyp.OutputAdressLength + sdataVari.Kartentyp.OutputAdressDiagnosticLength
                    Else
                        iAdressLegthInput = sdataVari.InputAdressSize - 1
                        iAdressLegthOutput = sdataVari.OutputAdressSize - 1
                        'KWS gibt alle Adresslängen mit Diagnose (+1 Byte) an, Hier bei IFM muss di eDiagnose abgezogen werden, nicht bei ET200SP IO-LINK
                        If iAdressLegthInput = 1 And iAdressLegthOutput <= 1 Then
                            iAdressVariInput = iAdressVariInput + 2
                            iAdressVariOutput = iAdressVariOutput + 1
                        ElseIf iAdressLegthInput <= 1 And iAdressLegthOutput = 1 Then 'beide null ist Fehler
                            iAdressVariInput = iAdressVariInput + 2
                            iAdressVariOutput = iAdressVariOutput + 1
                        ElseIf iAdressLegthInput <= 2 And iAdressLegthOutput <= 2 Then
                            sdataVari.OutputAdress = "AB" & iAdressVariOutput
                            iAdressVariInput = iAdressVariInput + 3
                            iAdressVariOutput = iAdressVariOutput + 2
                        ElseIf iAdressLegthInput <= 4 And iAdressLegthOutput <= 4 Then
                            iAdressVariInput = iAdressVariInput + 5
                            iAdressVariOutput = iAdressVariOutput + 4
                        ElseIf iAdressLegthInput <= 8 And iAdressLegthOutput <= 8 Then
                            iAdressVariInput = iAdressVariInput + 9
                            iAdressVariOutput = iAdressVariOutput + 8
                        ElseIf iAdressLegthInput <= 16 And iAdressLegthOutput <= 16 Then
                            iAdressVariInput = iAdressVariInput + 17
                            iAdressVariOutput = iAdressVariOutput + 16
                        ElseIf iAdressLegthInput <= 32 And iAdressLegthOutput <= 32 Then
                            iAdressVariInput = iAdressVariInput + 33
                            iAdressVariOutput = iAdressVariOutput + 32
                        Else
                            MsgBox "Fehler in der Adresslänge von: " & sData.KWSBMK
                        End If
                    End If
                    sdataVari.SubSlot = iSubSlot
                    iSubSlot = iSubSlot + 1
                    rData.Addobj sdataVari
                Next
                             
                ' Adressen für nächste Karte schreiben
                iInputStartAdress = iAdressVariInput
                iOutputStartAdress = iAdressVariOutput
            Else
                MsgBox "Adressen für Kartentyp: " & sdataVari.Kartentyp.Kartentyp & " kann nicht berechnet werden"
            End If
        Next i
        
    ElseIf Me.Item(1).Kartentyp.PLCtyp = "FESTO MPA" Then
        ' hier alles abhandeln was die Festo Ventilinsel betrifft
        'suchen der Datensätze mit den Grundmodulen
        For Each sDataFesto In Me
            If sDataFesto.Kartentyp.Kartentyp = "FESTO MPA1S" Then
                If sDataFesto.Kartentyp.Datatyp = "Byte" Then
                    If sDataFesto.Kartentyp.SignalDirection = "Input" Or sDataFesto.Kartentyp.SignalDirection = "InOut" Then
                        sDataFesto.Adress = "EB" & iInputStartAdress
                        sDataFesto.InputAdressPerSlot = iInputStartAdress
                        iInputStartAdress = iInputStartAdress + sDataFesto.Kartentyp.InputAdressLength + sDataFesto.Kartentyp.InputAdressDiagnosticLength
                    ElseIf sDataFesto.Kartentyp.SignalDirection = "Output" Then
                        sDataFesto.Adress = "AB" & iOutputStartAdress
                        sDataFesto.OutputAdressPerSlot = iOutputStartAdress
                        iOutputStartAdress = iOutputStartAdress + sDataFesto.Kartentyp.OutputAdressLength + sDataFesto.Kartentyp.OutputAdressDiagnosticLength
                    ElseIf sDataFesto.Kartentyp.SignalDirection = "InOut" Then
                         sDataFesto.OutputAdressPerSlot = iOutputStartAdress
                        iOutputStartAdress = iOutputStartAdress + sDataFesto.Kartentyp.OutputAdressLength + sDataFesto.Kartentyp.OutputAdressDiagnosticLength
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sDataFesto.Kartentyp.Kartentyp
                    End If
                End If
                rDataSearch.Addobj sDataFesto
                rData.Addobj sDataFesto
            End If
        Next
        For Each sDataFestoSearch In Me
            If Not (sDataFestoSearch.Kartentyp.Kartentyp = "FESTO MPA1S" Or sDataFestoSearch.Kartentyp.Kartentyp = "VMPA1-FB-SPU") Then
                
                For Each sDataFesto In rDataSearch
                    If sDataFestoSearch.Steckplatz = sDataFesto.Steckplatz Then 'Adresse des Steckplatzes suchen
                        'Testen auf welchem Subslot
                        iSubSlot = sDataFestoSearch.Kanal Mod 8
                        If iSubSlot = 1 Then
                            If sDataFestoSearch.Anschluss1 = "2" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".1"
                            Else
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".0"
                            End If
                        ElseIf iSubSlot = 3 Then
                            If sDataFestoSearch.Anschluss1 = "2" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".3"
                            Else
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".2"
                            End If
                        ElseIf iSubSlot = 5 Then
                            If sDataFestoSearch.Anschluss1 = "2" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".5"
                            Else
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".4"
                            End If
                        ElseIf iSubSlot = 7 Then
                            If sDataFestoSearch.Anschluss1 = "2" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".7"
                            Else
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".6"
                            End If
                        ElseIf iSubSlot = 2 Then ' z.B Kanal 50 CPX 2x3/2 mono
                            If sDataFestoSearch.Anschluss1 = "4" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".0"
                            End If
                        ElseIf iSubSlot = 4 Then ' z.B Kanal 52 CPX 2x3/2 mono
                            If sDataFestoSearch.Anschluss1 = "4" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".2"
                            End If
                        ElseIf iSubSlot = 6 Then ' z.B Kanal 54 CPX 2x3/2 mono
                            If sDataFestoSearch.Anschluss1 = "4" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".4"
                            End If
                        ElseIf iSubSlot = 0 Then ' z.B Kanal 56 CPX 2x3/2 mono
                            If sDataFestoSearch.Anschluss1 = "4" Then
                                sDataFestoSearch.Adress = "A" & sDataFesto.OutputAdressPerSlot & ".6"
                            End If
                        End If
                    End If
                Next
                rData.Addobj sDataFestoSearch
            End If
        Next
    ElseIf Me.Item(1).Kartentyp.PLCtyp = "Turck TBEN-S" Then
        ' hier alles für TURCK RFID
        For Each sDataTurck In Me
            sDataTurck.InputAdressPerSlot = iInputStartAdress
            sDataTurck.OutputAdressPerSlot = iOutputStartAdress
            If sDataTurck.Kartentyp.Kartentyp <> "TBEN-S IO" Then
                iInputStartAdress = iInputStartAdress + sDataTurck.Kartentyp.InputAdressLength + sDataTurck.Kartentyp.InputAdressDiagnosticLength
                iOutputStartAdress = iOutputStartAdress + sDataTurck.Kartentyp.OutputAdressLength + sDataTurck.Kartentyp.OutputAdressDiagnosticLength
            Else
                sDataTurck.Adress = vbNullString
            End If
            rData.Addobj sDataTurck
        Next
           
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
            Set SlotData = Me.searchKanalBelegungenSteckplatz(iStationsnummer, y)
            iAdressVariInput = 0
            iAdressVariOutput = 0
        
            For Each sData In SlotData
                sData.InputAdressPerSlot = iInputStartAdress
                sData.OutputAdressPerSlot = iOutputStartAdress
                If sData.Kartentyp.SameAdressRangeForOutput Then
                    ' prüfen was größer ist
                    If iInputStartAdress > iOutputStartAdress Then
                        sData.InputAdressPerSlot = iInputStartAdress
                        sData.OutputAdressPerSlot = sData.InputAdressPerSlot
                    Else
                        sData.OutputAdressPerSlot = iOutputStartAdress
                        sData.InputAdressPerSlot = sData.OutputAdressPerSlot
                    End If
                Else
                    ' hier normale Karten ohne extras
                    sData.InputAdressPerSlot = iInputStartAdress
                    sData.OutputAdressPerSlot = iOutputStartAdress
                End If
                If sData.Kartentyp.Datatyp = "Byte" Then
                    If sData.Kartentyp.SignalDirection = "Input" Then
                        sData.Adress = "EB" & iInputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                    ElseIf sData.Kartentyp.SignalDirection = "InOut" Then
                        sData.Adress = "EB" & iInputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                        sData.OutputAdress = "AB" & iOutputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                    ElseIf sData.Kartentyp.SignalDirection = "Output" Then
                        sData.Adress = "AB" & iOutputStartAdress + (sData.Kanal * sData.Kartentyp.OutputAdressLength)
                        sData.OutputAdress = sData.Adress
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sData.Kartentyp.Kartentyp
                    End If
                ElseIf sData.Kartentyp.Datatyp = "Word" Then
                    If sData.Kartentyp.SignalDirection = "Input" Then
                        sData.Adress = "EW" & iInputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                    ElseIf sData.Kartentyp.SignalDirection = "InOut" Then
                        sData.Adress = "EW" & iInputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                        sData.OutputAdress = "AW" & iOutputStartAdress + (sData.Kanal * sData.Kartentyp.InputAdressLength)
                    ElseIf sData.Kartentyp.SignalDirection = "Output" Then
                        sData.Adress = "AW" & iOutputStartAdress + (sData.Kanal * sData.Kartentyp.OutputAdressLength)
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sData.Kartentyp.Kartentyp
                    End If
                ElseIf sData.Kartentyp.Datatyp = "Bit" Then
                    'die Bit Berechnung
                    If sData.Kartentyp.SignalDirection = "Input" Then
                        If sData.Kanal <= 7 Then
                            sData.Adress = "E" & iInputStartAdress & "." & sData.Kanal
                        ElseIf sData.Kanal >= 8 And sData.Kanal <= 15 Then
                            sData.Adress = "E" & iInputStartAdress + 1 & "." & sData.Kanal - 8
                        ElseIf sData.Kanal >= 16 And sData.Kanal <= 23 Then
                            sData.Adress = "E" & iInputStartAdress + 2 & "." & sData.Kanal - 16
                        ElseIf sData.Kanal >= 24 And sData.Kanal <= 31 Then
                            sData.Adress = "E" & iInputStartAdress + 3 & "." & sData.Kanal - 24
                        End If
                    ElseIf sData.Kartentyp.SignalDirection = "Output" Then
                        If sData.Kanal <= 7 Then
                            sData.Adress = "A" & iOutputStartAdress & "." & sData.Kanal
                        ElseIf sData.Kanal >= 8 And sData.Kanal <= 15 Then
                            sData.Adress = "A" & iOutputStartAdress + 1 & "." & sData.Kanal - 8
                        ElseIf sData.Kanal >= 16 And sData.Kanal <= 23 Then
                            sData.Adress = "A" & iOutputStartAdress + 2 & "." & sData.Kanal - 16
                        ElseIf sData.Kanal >= 24 And sData.Kanal <= 31 Then
                            sData.Adress = "A" & iOutputStartAdress + 3 & "." & sData.Kanal - 24
                        End If
                    Else
                        MsgBox "Datentyp der Karte passt nicht: " & sData.Kartentyp.Kartentyp
                    End If
                Else
                    MsgBox "Adressen für Kartentyp: " & sData.Kartentyp.Kartentyp & " kann nicht berechnet werden"
                End If
                ' Ergebnis einsammeln
                rData.Addobj sData
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

Public Function correctFestoMPA(ByRef MPAconfig As cFestoMPA) As cKanalBelegungen
    ' Korrektur der Ventilinseldaten

    Dim sData As New cBelegung                   ' sData als neue Variable unter der Klasse cBelegung (sData = gespeicherte SPS-Belegungsdaten)
    Dim rData As New cKanalBelegungen
    Dim cData As New cKanalBelegungen
    Dim iSubSlot As Long
    Dim iSubAnschluss As Long
    
    'sort by channel
    Set cData = Me.SortChannel
    iSubAnschluss = 0

    'Startwerte setzen
    If MPAconfig.AnschlussplatteMPA = 0 Then
       MPAconfig.SteckplatzMPA = 1
       MPAconfig.KanalMPA = 1
    End If
    iSubSlot = 0
    
    For Each sData In cData
        If sData.Kartentyp.PLCtyp = "FESTO MPA" Then
            ' firstrun
            If MPAconfig.KanalMPA = 1 And iSubSlot = 0 Then
                MPAconfig.AnschlussplatteMPA = 1
                rData.Add "ADRESS", "FESTO_MPA", 0, sData.Stationsnummer, "FESTO MPA1S", MPAconfig.AnschlussplatteMPA
            End If
            '#### MPA Anschlussplatten ermitteln ####
            If MPAconfig.SteckplatzMPA > 4 Then
                MPAconfig.AnschlussplatteMPA = MPAconfig.AnschlussplatteMPA + 1
                If MPAconfig.AnschlussplatteMPA = 4 Or MPAconfig.AnschlussplatteMPA = 8 Or MPAconfig.AnschlussplatteMPA = 12 Or MPAconfig.AnschlussplatteMPA = 16 Or MPAconfig.AnschlussplatteMPA = 20 Then
                    rData.Add "ADRESS", "FESTO_MPA", 0, sData.Stationsnummer, "VMPA1-FB-SPU", MPAconfig.AnschlussplatteMPA 'Versorgungsplatte
                    MPAconfig.AnschlussplatteMPA = MPAconfig.AnschlussplatteMPA + 1
                    MPAconfig.ZwischeneinspeisungMPA = MPAconfig.ZwischeneinspeisungMPA + 1
                End If
                rData.Add "ADRESS", "FESTO_MPA", 0, sData.Stationsnummer, "FESTO MPA1S", MPAconfig.AnschlussplatteMPA 'Versorgungsplatte
                MPAconfig.SteckplatzMPA = 1
            End If
            ' diese Anschlussplatten sind nur zur Einspeisung
           
            sData.Steckplatz = MPAconfig.AnschlussplatteMPA
            sData.Kanal = MPAconfig.KanalMPA
            sData.SubSlot = MPAconfig.SteckplatzMPA
            'BMK bilden mit Kennbuchstaben KH
            sData.SPSBMK = "KH" & MPAconfig.SteckplatzMPA + ((MPAconfig.AnschlussplatteMPA - MPAconfig.ZwischeneinspeisungMPA) * 4) - 4
            
            If sData.Kartentyp.MaxChannelPerSlot = 0 Then
                MPAconfig.KanalMPA = MPAconfig.KanalMPA + 2
                MPAconfig.SteckplatzMPA = MPAconfig.SteckplatzMPA + 1
            ElseIf sData.Kartentyp.MaxChannelPerSlot = 1 Then
                If iSubSlot = 0 Then             ' Spule 12 ungerade Adressen
                    iSubSlot = iSubSlot + 1
                    MPAconfig.KanalMPA = MPAconfig.KanalMPA + 1
                    sData.Anschluss1 = 2
                    If Right(sData.Segmentvorlage, 1) = "0" Then
                        sData.Segmentvorlage = Replace(sData.Segmentvorlage, "KANAL_0", "KANAL_1")
                    End If
                Else                             ' Spule 14  gerade Adressen
                    iSubSlot = 0
                    MPAconfig.KanalMPA = MPAconfig.KanalMPA + 1
                    MPAconfig.SteckplatzMPA = MPAconfig.SteckplatzMPA + 1
                    sData.Anschluss1 = 4
                    If Right(sData.Segmentvorlage, 1) = "1" Then
                        sData.Segmentvorlage = Replace(sData.Segmentvorlage, "KANAL_1", "KANAL_0")
                    End If
                End If
            Else
                MsgBox "Fehler in den Daten, Kartentyp: " & sData.Kartentyp.Kartentyp
            End If
            ' Anschlüsse schreiben
            iSubAnschluss = MPAconfig.KanalMPA Mod 2
            If iSubAnschluss = 0 Then
                sData.Anschluss1 = 2
            Else
                sData.Anschluss1 = 4
            End If
        End If
        rData.AddDataSet sData
    Next

    Set correctFestoMPA = rData
End Function

Public Sub symbolischeAdressen()

    Dim sData As New cBelegung
    For Each sData In Me
        sData.SymbolischeAdresse = Trim$(sData.KWSBMK)
    Next
    
End Sub


