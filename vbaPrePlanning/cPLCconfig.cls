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
'controls access to it and ensures it will just contain people
' V0.1
' 14.02.2020
' class ist neu
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Konfig)


Private objPLCconfig As New Collection

'Public Kartentyp As String                'Key
'Public Kanal As Integer               'Secondary key

Private Sub Class_Initialize()
    Set objPLCconfig = New Collection
End Sub

Private Sub Class_Terminate()
    Set objPLCconfig = Nothing
End Sub


Public Sub Add(ByVal Stationsnummer As Integer, ByVal Steckplatz As Integer, ByVal Kartentyp As String, Optional ByVal FirstInputAdress As Long, Optional ByVal FirstOutputAdress As Long)
    'create a new Dataset and add to collection

    Dim p As New cPLCconfigData
    p.Stationsnummer = Stationsnummer
    p.Steckplatz = Steckplatz
    p.Kartentyp.Kartentyp = Kartentyp
    p.FirstInputAdress = FirstInputAdress
    p.FirstOutputAdress = FirstOutputAdress
    
    objPLCconfig.Add p
End Sub

Public Sub AddDataSet(ByRef dataSet As cPLCconfigData)
    'create a Dataset and add to collection

    Dim p As New cPLCconfigData
    p.key = dataSet.key
    p.Stationsnummer = dataSet.Stationsnummer
    p.Steckplatz = dataSet.Steckplatz
    p.Kartentyp = dataSet.Kartentyp
    p.FirstInputAdress = dataSet.FirstInputAdress
    p.FirstOutputAdress = dataSet.FirstOutputAdress
    
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

