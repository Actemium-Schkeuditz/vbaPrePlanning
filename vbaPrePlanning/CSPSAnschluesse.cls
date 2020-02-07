VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSPSAnschluesse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this collection could contain anything, but the class
'controls access to it and ensures it will just contain people


Private objSPSAnschluesse As New Collection

'Public Kartentyp As String                'Key
'Public Kanal As Integer               'Secondary key

Private Sub Class_Initialize()
    Set objSPSAnschluesse = New Collection
End Sub

Private Sub Class_Terminate()
    Set objSPSAnschluesse = Nothing
End Sub

Public Sub Add(Kartentyp As String, Kanal As Integer, Optional Anschluss1 As String, Optional Anschluss2 As String, Optional Anschluss3 As String, Optional Anschluss4 As String, Optional AnschlussVS As String, Optional AnschlussM As String)
    'create a new person and add to collection

    Dim p As New cAnschluss
    p.Kartentyp = Kartentyp
    p.Kanal = Kanal
    p.Anschluss_1 = Anschluss1
    p.Anschluss_2 = Anschluss2
    p.Anschluss_3 = Anschluss3
    p.Anschluss_4 = Anschluss4
    p.Anschluss_M = AnschlussM
    p.Anschluss_VS = AnschlussVS
    
    objSPSAnschluesse.Add p
End Sub

Public Sub AddDataSet(dataSet As cAnschluss)
    'create a new person and add to collection

    Dim p As New cAnschluss
    p.Kartentyp = dataSet.Kartentyp
    p.Kanal = dataSet.Kanal
    p.Anschluss_1 = dataSet.Anschluss_1
    p.Anschluss_2 = dataSet.Anschluss_2
    p.Anschluss_3 = dataSet.Anschluss_3
    p.Anschluss_4 = dataSet.Anschluss_4
    p.Anschluss_M = dataSet.Anschluss_M
    p.Anschluss_VS = dataSet.Anschluss_VS
    
    objSPSAnschluesse.Add p
End Sub

Public Property Get Count() As Long
    'return the number of Datasets
    Count = objSPSAnschluesse.Count
End Property

Public Sub Remove(NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objSPSAnschluesse.Remove NameOrNumber
End Sub

Public Property Get Item(Index As Variant) As cAnschluss
    'Attribute Item.VB_UserMemId = 0
    Set Item = objSPSAnschluesse.Item(Index)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objSPSAnschluesse.[_NewEnum]
End Property

Public Sub ReadExcelDataConnectionToCollection(dataColection As CSPSAnschluesse)
    'Reads the table on the Einbauorte sheet and adds
    'Conection points to the collection. It also adds values
    Dim rCell As Range
    Dim rTable As Range
    Dim sEinbauort As String
    Dim tablename As String

    On Error GoTo ErrorHandle
    tablename = "SPS_Anschluesse"

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
            If IsNumeric(.Value) Then
                sEinbauort = Str$(.Value)
            Else
                sEinbauort = .Value
            End If
            ' die gelesenen Daten an hängen
            Me.Add rCell.Offset(0, 0).Value, rCell.Offset(0, 1).Value, rCell.Offset(0, 2).Value, rCell.Offset(0, 3).Value, rCell.Offset(0, 4).Value, rCell.Offset(0, 5).Value, rCell.Offset(0, 7).Value, rCell.Offset(0, 6).Value ', rCell.Offset(0, 8).Value
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

Public Function searchAnschluss(pKartentyp As String, pKanal As Integer, data As CSPSAnschluesse) As CSPSAnschluesse
    ' Suchen nach den Anschlüssen passend zum Kartentyp und zum Kanal
    Dim sData As New cAnschluss
    'searchAnschluss = Nothing
    For Each sData In data
        If sData.Kartentyp = pKartentyp Then
            If sData.Kanal = pKanal Then
                Me.AddDataSet sData
                'Debug.Print searchAnschluss.Item(1).Kartentyp;
                Exit For                         ' nur einmal suchen dann beenden
            End If
        End If
    Next
    If Me.Count = 0 Then
        Me.Add vbNullString, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    'Test der Suche
    '  Debug.Print searchAnschluss.Item(1).Kartentyp; vbTab; searchAnschluss.Item(1).Kanal; vbTab; searchAnschluss.Item(1).Anschluss_1; vbTab; searchAnschluss.Item(1).Anschluss_2; vbTab; searchAnschluss.Item(1).Anschluss_3; vbTab; searchAnschluss.Item(1).Anschluss_4; vbTab; searchAnschluss.Item(1).Anschluss_M; vbTab; searchAnschluss.Item(1).Anschluss_VS
End Function

