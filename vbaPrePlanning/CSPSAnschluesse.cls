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
'@folder (Daten.SPS-Anschlüsse)

Private objSPSAnschluesse As New Collection

'Public Kartentyp As String                'Key
'Public Kanal As Integer               'Secondary key

Private Sub Class_Initialize()
    Set objSPSAnschluesse = New Collection
End Sub

Private Sub Class_Terminate()
    Set objSPSAnschluesse = Nothing
End Sub

Public Sub Add(ByVal Kartentyp As String, ByVal Kanal As Integer, Optional ByVal Anschluss1 As String, Optional ByVal Anschluss2 As String, Optional ByVal Anschluss3 As String, Optional ByVal Anschluss4 As String, Optional ByVal AnschlussVS As String, Optional ByVal AnschlussM As String)
    'create a new Dataset and add to collection

    Dim p As New cAnschluss
    p.Kartentyp.Kartentyp = Kartentyp
    p.Kanal = Kanal
    p.Anschluss1 = Anschluss1
    p.Anschluss2 = Anschluss2
    p.Anschluss3 = Anschluss3
    p.Anschluss4 = Anschluss4
    p.AnschlussM = AnschlussM
    p.AnschlussVS = AnschlussVS
    
    objSPSAnschluesse.Add p
End Sub

Public Sub AddDataSet(ByRef dataSet As cAnschluss)
    'create a Dataset and add to collection

    Dim p As New cAnschluss
    p.Kartentyp.Kartentyp = dataSet.Kartentyp.Kartentyp
    p.Kanal = dataSet.Kanal
    p.Anschluss1 = dataSet.Anschluss1
    p.Anschluss2 = dataSet.Anschluss2
    p.Anschluss3 = dataSet.Anschluss3
    p.Anschluss4 = dataSet.Anschluss4
    p.AnschlussM = dataSet.AnschlussM
    p.AnschlussVS = dataSet.AnschlussVS
    
    objSPSAnschluesse.Add p
End Sub

Public Property Get Count() As Long
    'return the number of Datasets
    Count = objSPSAnschluesse.Count
End Property

Public Sub Remove(ByRef NameOrNumber As Variant)
    'remove this Anschluesse from collection
    objSPSAnschluesse.Remove NameOrNumber
End Sub
'Attribute Item.VB_UserMemId = 0
'@DefaultMember
Public Property Get Item(ByRef Index As Variant) As cAnschluss
Attribute Item.VB_UserMemId = 0
    Set Item = objSPSAnschluesse.Item(Index)
End Property
'Attribute NewEnum.VB_UserMemId = -4
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = objSPSAnschluesse.[_NewEnum]
End Property

Public Sub ReadExcelDataConnectionToCollection(ByRef dataColection As CSPSAnschluesse)
    'Reads the table on the Einbauorte sheet and adds
    'Conection points to the collection. It also adds values
    Dim rCell As Range
    Dim rTable As Range
    Dim tablename As String

    On Error GoTo ErrorHandle
    tablename = "SPS_Anschluesse"

    'First we check if the table is intact. The
    'check is NOT complete.
    ThisWorkbook.Worksheets(tablename).Activate
    'Filter aus, aber nicht löschen
            If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
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

Public Function searchAnschluss(ByVal pKartentyp As String, ByVal pKanal As Integer, ByVal data As CSPSAnschluesse) As CSPSAnschluesse ' only ByVal works
    ' Suchen nach den Anschlüssen passend zum Kartentyp und zum Kanal
    Dim sData As New cAnschluss
    Dim sKartentyp As String
    Dim rData As New CSPSAnschluesse
    Set searchAnschluss = Nothing
    
    For Each sData In data
         
        If sData.Kartentyp.Kartentyp = pKartentyp Then
            sKartentyp = sData.Kartentyp.Kartentyp
            If sData.Kanal = pKanal Then
                rData.AddDataSet sData
                'Debug.Print searchAnschluss.Item(1).Kartentyp;
                Exit For                         ' nur einmal suchen dann beenden
            End If
        End If
    Next
    If rData.Count = 0 Then
        rData.Add sKartentyp, 0, "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER", "FEHLER"
    End If
    
    Set searchAnschluss = rData
    'Test der Suche
    '  Debug.Print searchAnschluss.Item(1).Kartentyp; vbTab; searchAnschluss.Item(1).Kanal; vbTab; searchAnschluss.Item(1).Anschluss_1; vbTab; searchAnschluss.Item(1).Anschluss_2; vbTab; searchAnschluss.Item(1).Anschluss_3; vbTab; searchAnschluss.Item(1).Anschluss_4; vbTab; searchAnschluss.Item(1).Anschluss_M; vbTab; searchAnschluss.Item(1).Anschluss_VS
End Function





