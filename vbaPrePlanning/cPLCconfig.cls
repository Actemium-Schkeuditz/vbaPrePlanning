VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPLCconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this collection could contain anything, but the class
'controls access to it and ensures it will just contain cPLCconfigData
' V0.3
' 25.02.2020
' new Propertys
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Konfig)


Private objPLCconfig As New Collection

'Public Kartentyp As String                'Key

Private Sub Class_Initialize()
    Set objPLCconfig = New Collection
End Sub

Private Sub Class_Terminate()
    Set objPLCconfig = Nothing
End Sub

Public Sub Add(ByVal Stationsnummer As Integer, ByVal Steckplatz As Integer, ByVal Kartentyp As String, Optional ByVal Key As String, Optional ByVal FirstInputAdress As Long, Optional ByVal FirstOutputAdress As Long, Optional ByVal ReserveChannelsBefor As Long, Optional ByVal ReserveChannelsAfter As Long, Optional ByVal ReserveChannel As Long, Optional ByVal ReserveSlot As Long)
    'create a new Dataset and add to collection

    Dim p As New cPLCconfigData
    p.Key = Key
    p.Stationsnummer = Stationsnummer
    p.Steckplatz = Steckplatz
    p.Kartentyp.Kartentyp = Kartentyp
    p.FirstInputAdress = FirstInputAdress
    p.FirstOutputAdress = FirstOutputAdress
    p.ReserveChannelsBefor = ReserveChannelsBefor
    p.ReserveChannelsAfter = ReserveChannelsAfter
    p.ReserveChannelPerSlot = ReserveChannel
    p.ReserveSlot = ReserveSlot
    
    objPLCconfig.Add p
End Sub

Public Sub AddDataSet(ByRef dataSet As cPLCconfigData)
    'create a Dataset and add to collection

    Dim p As New cPLCconfigData
    p.Key = dataSet.Key
    p.Stationsnummer = dataSet.Stationsnummer
    p.Steckplatz = dataSet.Steckplatz
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.FirstInputAdress = dataSet.FirstInputAdress
    p.FirstOutputAdress = dataSet.FirstOutputAdress
    p.ReserveChannelsBefor = dataSet.ReserveChannelsBefor
    p.ReserveChannelsAfter = dataSet.ReserveChannelsAfter
    p.ReserveChannelPerSlot = dataSet.ReserveChannelPerSlot
    p.ReserveSlot = dataSet.ReserveSlot
    
    objPLCconfig.Add p
End Sub

Public Sub Addobj(ByRef obj As cPLCconfigData)
    objPLCconfig.Add obj
End Sub

Public Property Get Count() As Long
    'return the number of Datasets
    Count = objPLCconfig.Count
End Property

Public Sub Remove(ByRef NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objPLCconfig.Remove NameOrNumber
End Sub

'Attribute Item.VB_UserMemId = 0
'@DefaultMember
Public Property Get Item(ByRef Index As Variant) As cPLCconfigData
Attribute Item.VB_UserMemId = 0
    Set Item = objPLCconfig.Item(Index)
End Property

'Attribute NewEnum.VB_UserMemId = -4
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objPLCconfig.[_NewEnum]
End Property

Public Sub Append(ByRef data As cPLCconfig)

    Dim sData As New cPLCconfigData
    For Each sData In data
        Me.Addobj sData
    Next
End Sub

Public Function Sort() As cPLCconfig
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim bln As Boolean
    Dim lngCount As Long
    Dim arr() As Long
    Dim ppl As cPLCconfig
    
    lngCount = Me.Count
    If lngCount > 0 Then
        ReDim arr(0 To lngCount - 1)
        For i = 0 To lngCount - 1: arr(i) = i + 1: Next

        For i = 1 To lngCount - 1
            k = arr(i)
            j = i - 1
            bln = False
            Do
            
                If Me.Item(arr(j)).ComparePLCconfigTo(Me.Item(k)) > 0 Then
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

    Set ppl = New cPLCconfig
    For i = 0 To lngCount - 1: ppl.Addobj Me.Item(arr(i)): Next

    Set Sort = ppl
End Function

Public Sub ReadPLCConfigData(ByRef tablename As String) ', ByRef dataColection As cPLCconfig)
    'Reads the table on the "EplSheet" sheet and adds
    'Chanels to the collection. It also adds values
    Dim rCell As Range
    Dim rTable As Range
    Dim ws1 As Worksheet
    
    On Error GoTo ErrorHandle

    With ThisWorkbook
        ' alte Daten löschen
        'first check if table exist
        If WorksheetExist((tablename), ws1) = True Then
           
            'Second we check if the table is intact. The
            'check is NOT complete.
            ThisWorkbook.Worksheets(tablename).Activate
            'Filter aus, aber nicht löschen
            If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
             Application.ScreenUpdating = False
            If Len(Range("A2")) = 0 Then
                MsgBox "The table must start in cell A2. Adding products has been aborted.", vbCritical, "Error"
                'Set dataColection = Nothing
                Exit Sub
            End If

            Set rTable = Range("A2")             ' Auslesen ohne Überschrift

            With rTable
                If Len(.Offset(0, 0).Value) = 0 Or Len(.Offset(0, 1).Value) = 0 Then
                    MsgBox "The products table has been destoyed and there is no or too few data.", vbCritical, "Error"
                    'Set dataColection = Nothing
                    Exit Sub
                End If
            End With

            'If the table is okay, we start reading data.
            'The range is expanded to include the last
            'cell with content.
            'If Len(rTable.Offset(2, 0).Formula) = 0 Then
            '   Set rTable = rTable.Offset(0, 9)
            'Else
            Set rTable = Range(rTable.Offset(0, 0), rTable.End(xlDown))
            'End If

            For Each rCell In rTable
                With rCell
                    ' prüfen ob Stationsnummer nicht leer
                    If rCell.Offset(0, 0).Value <> vbNullString Then
                        '      sEinbauort = str$(.Value)
                        ' die gelesenen Daten anhängen
                        'Datensatz 1
                        Me.Add CInt(rCell.Offset(0, 0).Value), CInt(rCell.Offset(0, 1).Value), rCell.Offset(0, 2).Value, 0, CInt(rCell.Offset(0, 3).Value), CInt(rCell.Offset(0, 4).Value), CInt(rCell.Offset(0, 5).Value), CInt(rCell.Offset(0, 6).Value), CInt(rCell.Offset(0, 7).Value), CInt(rCell.Offset(0, 8).Value)
                    End If
                End With
            Next
            'MsgBox "gelesen"
        End If
    End With
BeforeExit:
    Set rCell = Nothing
    Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim einlesen.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Sub writePLCConfigToExcel(ByVal tablename As String)
    Dim wdata As New cPLCconfigData
    Dim ws1 As Worksheet
    Dim iSpalteKanal As Long
    Dim i As Long
    Dim zeilenanzahl As Long
    
     On Error GoTo ErrorHandle
    
        iSpalteKanal = 0
        zeilenanzahl = 0
    With ThisWorkbook
        ' alte Daten löschen
        Application.ScreenUpdating = False
        If WorksheetExist((tablename), ws1) = True Then
            Application.DisplayAlerts = False
            ' löschen deaktiviert
            '.Sheets(tablename).Delete
            Application.DisplayAlerts = True
        Else
            'Worksheets anlegen
            .Sheets.Add after:=Sheets(Worksheets.Count)
            .ActiveSheet.Name = tablename
        End If
        ' Daten einschreiben
      
        ThisWorkbook.Worksheets(tablename).Activate
     
        ' Tabellenkopf
        .ActiveSheet.Cells(1, 1) = "Stationsnummer"
        .ActiveSheet.Cells(1, 2) = "Steckplatz"
        .ActiveSheet.Cells(1, 3) = "Kartentyp"
        .ActiveSheet.Cells(1, 4) = "Eingangsadresse"
        .ActiveSheet.Cells(1, 5) = "Ausgangsadresse"
        .ActiveSheet.Cells(1, 6) = "Kanäle vor erstem Steckplatz"
        .ActiveSheet.Cells(1, 7) = "Kanäle nach Belegung"
        .ActiveSheet.Cells(1, 8) = "Reservekanäle pro Karte"
        .ActiveSheet.Cells(1, 9) = "Reserve Steckplatz"
     
        i = 2
        For Each wdata In Me
            .ActiveSheet.Cells(i, 1) = wdata.Stationsnummer
            .ActiveSheet.Cells(i, 2) = wdata.Steckplatz
            .ActiveSheet.Cells(i, 3) = wdata.Kartentyp.Kartentyp
            .ActiveSheet.Cells(i, 4) = wdata.FirstInputAdress
            .ActiveSheet.Cells(i, 5) = wdata.FirstOutputAdress
            i = i + 1
        Next
        i = 0
   
    End With


BeforeExit:
    'Set rCell = Nothing
    'Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler beim einlesen.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Function returnDatasetPerSlottyp(ByVal Station As Integer, ByVal sKartentyp As String) As cPLCconfig
    ' returns first Dataset for Station and plctyp
    On Error GoTo ErrorHandle
    Dim sData As New cPLCconfigData
    Dim rData As New cPLCconfig
    
    For Each sData In Me
        If sData.Stationsnummer = Station And sData.Kartentyp.Kartentyp = sKartentyp Then
            rData.Addobj sData
            Exit For
        End If
    
    Next
    
    
    If rData.Count = 0 Then
        rData.Add Station, 0, sKartentyp
    End If
    
    Set returnDatasetPerSlottyp = rData
    Exit Function
ErrorHandle:
    If Err.Number > 0 Then
        MsgBox "Fehler in den PLS Konfigurationen für Stationsnummer: " & Station
        Err.Clear
        Resume Next
    End If

End Function
Public Function returnDatasetPerStation(ByVal Station As Integer) As cPLCconfig
    ' returns all datasets for Station
    On Error GoTo ErrorHandle
    Dim sData As New cPLCconfigData
    Dim rData As New cPLCconfig
    
    For Each sData In Me
        If sData.Stationsnummer = Station Then
            rData.Addobj sData
        End If
    
    Next
    
    
    If rData.Count = 0 Then
        rData.Add Station, 0, ""
    End If
    
    Set returnDatasetPerStation = rData
    Exit Function
ErrorHandle:
    If Err.Number > 0 Then
        MsgBox "Fehler in den PLS Konfigurationen für Stationsnummer: " & Station
        Err.Clear
        Resume Next
    End If

End Function

Public Function returnFirstInputAdressePLCStation(ByVal Station As Integer) As Long
    ' Rückgabe der ersten Eingangsadressen für eine gesamte Station
    Dim sData As New cPLCconfigData
    Dim rData As Long
    For Each sData In Me
        If sData.Stationsnummer = Station Then
            If sData.FirstInputAdress > 0 Then
                rData = sData.FirstInputAdress
                Exit For
            End If
        End If
    Next
    
    returnFirstInputAdressePLCStation = rData
 
End Function

Public Function returnFirstOutputAdressePLCStation(ByVal Stationsnummer As Integer) As Long
    ' Rückgabe der ersten Ausgangsadressen für eine gesamte Station
    Dim sData As New cPLCconfigData
    Dim rData As Long
    
    For Each sData In Me
        If sData.Stationsnummer = Stationsnummer Then
            If sData.FirstOutputAdress > 0 Then
                rData = sData.FirstOutputAdress
                Exit For
            End If
        End If
    Next
    
    returnFirstOutputAdressePLCStation = rData
 
End Function

Public Function sumAdressesPerSlot(ByVal Station As Integer, ByVal Kanaldaten As cKanalBelegungen) As cPLCconfig

    Dim InputAdress As Long
    Dim OutputAdress As Long
    Dim iStartInputAdress As Long
    Dim iStartOutputAdress As Long
    Dim iTmpInputAdress As Long
    Dim iTmpOutputAdress As Long
    Dim iStartSteckplatz As Long
    Dim iSteckplatzdifferenz As Long
    Dim sData As New cBelegung
    Dim rData As New cPLCconfig
    Dim sPLCData As New cPLCconfigData
    Dim sPLCDataAdress As New cPLCconfig
    Dim sPLCDataAdressPerSlot As New cPLCconfigData
    Dim sPLCassignmentsPerSlot As New cKanalBelegungen
    Dim sDataPLCSlot As New cBelegung
      
    ' Startwerte setzen
    iStartInputAdress = Me.Item(1).FirstInputAdress
    iStartOutputAdress = Me.Item(1).FirstOutputAdress
    iStartSteckplatz = Me.Item(1).Steckplatz


    ' berechnen der Adressen für jeden Steckplatz
    For Each sPLCData In Me
        If sPLCData.Kartentyp.VariableAdressRange = False Then
            If sPLCData.Steckplatz > iStartSteckplatz Then
                ' Berechnung wenn sich Steckplatz ändert
                iSteckplatzdifferenz = sData.Steckplatz - iStartSteckplatz
                iStartSteckplatz = sData.Steckplatz
                InputAdress = iStartInputAdress + InputAdress + (iSteckplatzdifferenz * (sPLCData.Kartentyp.InputAdressLength + sPLCData.Kartentyp.InputAdressDiagnosticLength))
                OutputAdress = iStartOutputAdress + OutputAdress + (iSteckplatzdifferenz * (sPLCData.Kartentyp.OutputAdressLength + sPLCData.Kartentyp.OutputAdressDiagnosticLength))
            End If
            'sPLCData.FirstInputAdress = InputAdress
            'sPLCData.FirstOutputAdress = OutputAdress
            'sPLCDataAdress.Addobj sPLCData
        Else
            'Variable Adressvergabe
            'Für jede Karte einzeln berechnen
            'Suchen nach den Belegungen der Karten mit variabler Länge
            Set sPLCassignmentsPerSlot = Kanaldaten.searchKanalBelegungenSteckplatz(sPLCData.Stationsnummer, sPLCData.Steckplatz)
            sPLCassignmentsPerSlot.SortChannel
            iTmpInputAdress = 0
            iTmpOutputAdress = 0
            If sPLCData.Kartentyp.Kartentyp = "ET200SP 4IO-LINK" Then
                For Each sDataPLCSlot In sPLCassignmentsPerSlot
                    ' zusammenzählen der Adressbereiche der Sensoren
                    iTmpInputAdress = iTmpInputAdress + sDataPLCSlot.InputAdressSize
                    iTmpOutputAdress = iTmpOutputAdress + sDataPLCSlot.OutputAdressSize
                Next
                If iTmpInputAdress < 16 And iTmpOutputAdress < 16 Then
                    InputAdress = iStartInputAdress + InputAdress + 16
                    OutputAdress = iStartOutputAdress + OutputAdress + 16
                ElseIf iTmpInputAdress < 32 And iTmpOutputAdress < 32 Then
                    InputAdress = iStartInputAdress + InputAdress + 32
                    OutputAdress = iStartOutputAdress + OutputAdress + 32
                ElseIf iTmpInputAdress < 64 And iTmpOutputAdress < 64 Then
                    InputAdress = iStartInputAdress + InputAdress + 64
                    OutputAdress = iStartOutputAdress + OutputAdress + 64
                ElseIf iTmpInputAdress <= 126 And iTmpOutputAdress <= 128 Then
                    InputAdress = iStartInputAdress + InputAdress + 144
                    OutputAdress = iStartOutputAdress + OutputAdress + 128
                Else
                    MsgBox "Fehler Adressbereich des Steckplatzes" & sDataPLCSlot.Steckplatz & " auf Station: " & sDataPLCSlot.Stationsnummer
                End If
            ElseIf sPLCData.Kartentyp.Kartentyp = "AL1400" Then
                'todo Adressierung der IO-LINK master funktioniert nicht
                
                InputAdress = iStartInputAdress
                OutputAdress = iStartOutputAdress
            ElseIf sPLCData.Kartentyp.Kartentyp = "AL1402" Then
                'todo Adressierung der IO-LINK Master funktioniert nicht
                
                InputAdress = iStartInputAdress
                OutputAdress = iStartOutputAdress
            
            
            ElseIf sPLCData.Kartentyp.Kartentyp = "IFM IO-LINK" Then
                MsgBox "Fehler in Datensatz, IFM IO-Link als Kartentyp nicht zulässig, Station: " & sPLCData.Stationsnummer
                
                 
            End If
        End If
        sPLCData.Kartentyp.FirstInputAdress = InputAdress
        sPLCData.Kartentyp.FirstOutputAdress = OutputAdress
        sPLCDataAdress.Addobj sPLCData
    Next


    Set sumAdressesPerSlot = rData

End Function

Public Function ConfigPLCToDataset(Stationdata As cKanalBelegungen) As cPLCconfig
    'extract data from a Class cKanalbelegungen to cPLCconfig
    'writes PLC Config to Dataset
    Dim i As Long
    Dim OffsetSlot As Integer

    ' Class einbinden
    Dim sData As New cBelegung
    Dim dataConfig As New cPLCconfig
    Dim dataConfigSort As cPLCconfig
    Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchConfig As New cKanalBelegungen
    
    
   
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    'Set iStation = dataKanaele.returnStation
    Set iStation = Stationdata.returnStation
    '##### Suche nach allen verwendeten Kartentypen
    Dim iKartentyp As Collection
   
    
    '####### bearbeiten der Daten #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Dim pKartentyp As Variant
    Set dataConfig = Nothing
    'Dim iAdressOutput As Long
    'Dim iAdressInput As Long
    
    For Each pStation In iStation
        ' suchen der Datensätze pro Station
        Set dataSearchStation = Stationdata.searchDatasetPerStation(pStation)
        Set dataSearchConfig = dataSearchStation.returnAllSlotsPerRack
        ' Übertragen der Daten
        For Each sData In dataSearchConfig
            dataConfig.Add sData.Stationsnummer, sData.Steckplatz, sData.Kartentyp.Kartentyp, sData.Key, sData.InputAdressPerSlot, sData.OutputAdressPerSlot
        Next
    
        ' Sortieren der Steckplätze
        Set dataConfigSort = dataConfig.Sort
          
    Next
    ' Rückgabe des Ergebnises
    Set ConfigPLCToDataset = dataConfigSort
    
End Function

