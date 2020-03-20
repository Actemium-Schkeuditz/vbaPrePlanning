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

Private Sub Class_Initialize()
    Set objPLCconfig = New Collection
End Sub

Private Sub Class_Terminate()
    Set objPLCconfig = Nothing
End Sub

Public Sub Add(ByVal Stationsnummer As Integer, ByVal Steckplatz As Integer, ByVal Kartentyp As String, Optional ByVal Key As String, Optional ByVal FirstInputAdress As Long, Optional ByVal FirstOutputAdress As Long, Optional ByVal ReserveChannelsBefor As Long, Optional ByVal ReserveChannelsAfter As Long, Optional ByVal ReserveChannel As Long, Optional ByVal ReserveSlot As Long, Optional ByVal SubSlot As Integer, Optional ByVal UsedChannels As Long, Optional ByVal InputAdress As String, Optional ByVal OutputAdress As String)
    'create a new Dataset and add to collection

    Dim p As New cPLCconfigData
    p.Key = Key
    p.Stationsnummer = Stationsnummer
    p.Steckplatz = Steckplatz
    p.SubSlot = SubSlot
    p.Kartentyp.Kartentyp = Kartentyp
    p.FirstInputAdress = FirstInputAdress
    p.FirstOutputAdress = FirstOutputAdress
    p.ReserveChannelsBefor = ReserveChannelsBefor
    p.ReserveChannelsAfter = ReserveChannelsAfter
    p.ReserveChannelPerSlot = ReserveChannel
    p.ReserveSlot = ReserveSlot
    p.UsedChannels = UsedChannels
    p.InputAdress = InputAdress
    p.OutputAdress = OutputAdress
    
    objPLCconfig.Add p
End Sub

Public Sub AddDataSet(ByRef dataSet As cPLCconfigData)
    'create a Dataset and add to collection

    Dim p As New cPLCconfigData
    p.Key = dataSet.Key
    p.Stationsnummer = dataSet.Stationsnummer
    p.Steckplatz = dataSet.Steckplatz
    p.SubSlot = dataSet.SubSlot
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.FirstInputAdress = dataSet.FirstInputAdress
    p.FirstOutputAdress = dataSet.FirstOutputAdress
    p.ReserveChannelsBefor = dataSet.ReserveChannelsBefor
    p.ReserveChannelsAfter = dataSet.ReserveChannelsAfter
    p.ReserveChannelPerSlot = dataSet.ReserveChannelPerSlot
    p.ReserveSlot = dataSet.ReserveSlot
    p.UsedChannels = dataSet.UsedChannels
    p.InputAdress = dataSet.InputAdress
    p.OutputAdress = dataSet.OutputAdress
    
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

Public Sub writePLCConfigToExcel(ByVal tablename As String, ByVal PLCTyp As String)
    Dim wdata As New cPLCconfigData
    Dim ws1 As Worksheet

    Dim i As Long
    
    On Error GoTo ErrorHandle

    With ThisWorkbook
        ' alte Daten löschen
        Application.ScreenUpdating = False
        If WorksheetExist((tablename), ws1) = True Then
            Application.DisplayAlerts = False
            ' löschen deaktiviert
            .Sheets.[_Default](tablename).Delete
            Application.DisplayAlerts = True
            'Worksheets anlegen
            .Sheets.Add after:=Sheets.[_Default](Worksheets.Count)
            .ActiveSheet.Name = tablename
        Else
            'Worksheets anlegen
            .Sheets.Add after:=Sheets.[_Default](Worksheets.Count)
            .ActiveSheet.Name = tablename
        End If
        ' Daten einschreiben
      
        If PLCTyp = "ET200SP" Then
            .ActiveSheet.Tab.ColorIndex = 28
        ElseIf PLCTyp = "FESTO CPX" Or PLCTyp = "FESTO MPA" Then
            .ActiveSheet.Tab.ColorIndex = 32
        ElseIf PLCTyp = "IFM IO-LINK" Then
            .ActiveSheet.Tab.ColorIndex = 46
        ElseIf PLCTyp = "FU" Then
            .ActiveSheet.Tab.ColorIndex = 50
        Else
            .ActiveSheet.Tab.ColorIndex = 17
        End If
            
        ThisWorkbook.Worksheets.[_Default](tablename).Activate
     
        ' Tabellenkopf
        .ActiveSheet.Cells(1, 1) = "Stationsnummer"
        .ActiveSheet.Cells(1, 2) = "Steckplatz"
        .ActiveSheet.Cells(1, 3) = "Unter Steckplatz"
        .ActiveSheet.Cells(1, 4) = "Kartentyp"
        .ActiveSheet.Cells(1, 5) = "Eingangsadresse Start"
        .ActiveSheet.Cells(1, 6) = "Ausgangsadresse Start"
        .ActiveSheet.Cells(1, 7) = "freie Kanäle"
     
        i = 2
        For Each wdata In Me
            .ActiveSheet.Cells(i, 1) = wdata.Stationsnummer
            .ActiveSheet.Cells(i, 2) = wdata.Steckplatz
            .ActiveSheet.Cells(i, 3) = wdata.SubSlot
            .ActiveSheet.Cells(i, 4) = wdata.Kartentyp.Kartentyp
            If wdata.Kartentyp.PLCTyp = "IFM IO-LINK MASTER" Then 'And wdata.SubSlot > 0 Then
                If wdata.SubSlot = 1 Then
                    .ActiveSheet.Cells(i, 3) = 0
                    .ActiveSheet.Cells(i, 5) = wdata.FirstInputAdress
                    .ActiveSheet.Cells(i, 6) = wdata.FirstOutputAdress
                    .ActiveSheet.Cells(i, 7) = wdata.Kartentyp.MaxChannelPerSlot - wdata.UsedChannels
                    i = i + 1
                    .ActiveSheet.Cells(i, 1) = wdata.Stationsnummer
                    .ActiveSheet.Cells(i, 2) = wdata.Steckplatz
                    .ActiveSheet.Cells(i, 3) = wdata.SubSlot
                    .ActiveSheet.Cells(i, 4) = wdata.Kartentyp.Kartentyp
                End If
                .ActiveSheet.Cells(i, 5) = wdata.InputAdress
                .ActiveSheet.Cells(i, 6) = wdata.OutputAdress
            Else
                .ActiveSheet.Cells(i, 5) = wdata.FirstInputAdress
                .ActiveSheet.Cells(i, 6) = wdata.FirstOutputAdress
            End If
            If wdata.Kartentyp.PLCTyp <> "FESTO MPA" And wdata.Kartentyp.PLCTyp <> "IFM IO-LINK MASTER" Then
                .ActiveSheet.Cells(i, 7) = wdata.Kartentyp.MaxChannelPerSlot - wdata.UsedChannels
            Else
                .ActiveSheet.Cells(i, 7) = 0
            End If
            
            i = i + 1
        Next
        i = 0
   
    End With

BeforeExit:
    Set wdata = Nothing
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
        rData.Add Station, 0, vbNullString
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

Public Function ConfigPLCToDataset(ByRef Stationdata As cKanalBelegungen) As cPLCconfig
    'extract data from a Class cKanalbelegungen to cPLCconfig
    'writes PLC Config to Dataset
    Dim iLastUsedChannel As Long
    ' Class einbinden
    Dim sData As New cBelegung
    Dim sDataKanal As New cBelegung
    Dim dataConfig As New cPLCconfig
    Dim dataConfigSort As cPLCconfig
    'Dim dataKanaele As New cKanalBelegungen
    Dim dataSearchStation As New cKanalBelegungen
    Dim dataSearchConfig As New cKanalBelegungen
    
       
    '##### Suche nach allen Stationsnummern
    Dim iStation As Collection
    Set iStation = Stationdata.returnStation
    
    '####### bearbeiten der Daten #######
    ' Durchlauf für jede Station einzeln
    Dim pStation As Variant
    Set dataConfig = Nothing

    
    For Each pStation In iStation
        ' suchen der Datensätze pro Station
        Set dataSearchStation = Stationdata.searchDatasetPerStation(pStation)
        If dataSearchStation.Item(1).Kartentyp.PLCTyp = "FESTO MPA" Or dataSearchStation.Item(1).Kartentyp.PLCTyp = "IFM IO-LINK MASTER" Then
            Set dataSearchConfig = dataSearchStation
        Else
            Set dataSearchConfig = dataSearchStation.returnAllSlotsPerRack
            
        End If
        
        ' Übertragen der Daten
        For Each sData In dataSearchConfig
            iLastUsedChannel = -1
            If sData.Kartentyp.PLCTyp = "FESTO MPA" Or (sData.Kartentyp.PLCTyp = "IFM IO-LINK MASTER" And sData.Kanal <> 0) Then
                iLastUsedChannel = 0
            Else
                For Each sDataKanal In Stationdata
                    If sData.Stationsnummer = sDataKanal.Stationsnummer Then
                        If sData.Steckplatz = sDataKanal.Steckplatz Then
                            If sDataKanal.Kanal >= iLastUsedChannel Then
                                iLastUsedChannel = sDataKanal.Kanal
                            End If
                        End If
                    End If
                Next
            End If
            dataConfig.Add sData.Stationsnummer, sData.Steckplatz, sData.Kartentyp.Kartentyp, sData.Key, sData.InputAdressPerSlot, sData.OutputAdressPerSlot, 0, 0, 0, 0, sData.SubSlot, iLastUsedChannel, sData.Adress, sData.OutputAdress
        Next
    
        ' Sortieren der Steckplätze
        Set dataConfigSort = dataConfig.Sort
          
    Next
    ' Rückgabe des Ergebnises
    Set ConfigPLCToDataset = dataConfigSort
    
End Function





