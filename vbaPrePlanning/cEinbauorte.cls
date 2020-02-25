VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEinbauorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Coolection / class zum speichern der Einbauorte
' V0.3
' 20.02.2020
' new property "BMK"
' Christian Langrock
' christian.langrock@actemium.de

Option Explicit
'@folder(Kennzeichen.Einbauorte)

Private objEinbauorte As New Collection


Private Sub Class_Initialize()
    Set objEinbauorte = New Collection
End Sub

Private Sub Class_Terminate()
    Set objEinbauorte = Nothing
End Sub

Public Sub Add(ByRef Stationsnummer As Integer, Optional ByRef Einbauort As String, Optional ByRef Geraetetyp As String, Optional ByRef BMK As String)
    'create a new Dataset Einbauorte and add to collection

    Dim p As New cEinbauort
    p.Stationsnummer = Stationsnummer
    p.Einbauort = Einbauort
    p.Geraetetyp = Geraetetyp
    p.BMK = BMK
    
    objEinbauorte.Add p
End Sub

Public Sub AddDataSet(ByRef dataSet As cEinbauort)
    'create a Dataset and add to collection

    Dim p As New cEinbauort
    p.Stationsnummer = dataSet.Stationsnummer
    p.Einbauort = dataSet.Einbauort
    p.Geraetetyp = dataSet.Geraetetyp
    p.BMK = dataSet.BMK
   
    objEinbauorte.Add p
End Sub

Public Sub Addobj(ByRef obj As cEinbauorte)
    objEinbauorte.Add obj
End Sub

Public Property Get Count() As Long
    'return the number of Datasets
    Count = objEinbauorte.Count
End Property

Public Sub Remove(ByRef NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objEinbauorte.Remove NameOrNumber
End Sub

'Attribute Item.VB_UserMemId = 0
Public Property Get Item(ByRef Index As Variant) As cEinbauort
    Set Item = objEinbauorte.Item(Index)
End Property

'VBA BUG falls die Suchfunktion nicht funktioniert folgendes hier einfügen:
'Attribute NewEnum.VB_UserMemId = -4
'danach im Notepad die exortierte Klasse bearbeiten, die Auskommentierung für: "' Attribute NewEnum.VB_UserMemId = -4" entfernen und wieder importieren
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objEinbauorte.[_NewEnum]
End Property

Public Sub ReadExcelDataToCollection(ByRef tablename As String, ByRef dataColection As cEinbauorte)
    'Reads the table on the Einbauorte sheet and adds
    'Einbauorte to the collection. It also adds values
    'to the product's properties.
    Dim rCell As Range
    Dim rTable As Range
    'Dim sEinbauort As String

    On Error GoTo ErrorHandle

    'First we check if the table is intact. The
    'check is NOT complete.
    ThisWorkbook.Worksheets(tablename).Activate
    If Len(Range("A1")) = 0 Then
        MsgBox "The table must start in cell A1. Adding products has been aborted.", vbCritical, "Error"
        'Set EinbauorteData = Nothing
        Set dataColection = Nothing
        Exit Sub
    End If

    Set rTable = Range("A1")

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
            'If IsNumeric(.Value) Then
             '   sEinbauort = str$(.Value)
            'Else
            '    sEinbauort = .Value
            'End If
            ' die gelesenen Daten anhängen
            Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, rCell.Offset(0, 2).Value, rCell.Offset(0, 3).Value
  
        End With
    Next

BeforeExit:
    Set rCell = Nothing
    Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Procedure MakeProductCollection.", vbCritical, "Error"
    Resume BeforeExit
End Sub

Public Function searchEinbauort(ByVal pStationsnumer As Integer, ByVal data As cEinbauorte) As String
    ' Suchen nach dem Einbauort passend zur Stationsnummer
    Dim sData As cEinbauort
    searchEinbauort = vbNullString
    For Each sData In data
        If sData.Stationsnummer = pStationsnumer Then
            '        Debug.Print Data.Einbauort; vbTab; Data.Stationsnummer
            searchEinbauort = sData.Einbauort
            Exit For                             ' nur einmal suchen dann beenden
        End If
    Next
End Function

Public Function searchEinbauortDataset(ByVal pStationsnumer As Integer) As cEinbauorte
    ' Suchen nach dem Einbauort passend zur Stationsnummer
     Dim sData As cEinbauort
    Dim rData As New cEinbauorte
    For Each sData In Me
        If sData.Stationsnummer = pStationsnumer Then
            '        Debug.Print Data.Einbauort; vbTab; Data.Stationsnummer
            rData.AddDataSet sData
            Exit For                             ' nur einmal suchen dann beenden
        End If
    Next
    Set searchEinbauortDataset = rData
End Function

