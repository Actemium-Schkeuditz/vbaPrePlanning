VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBelegung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der SPS Belegungsdaten
' V0.1
' 10.02.2020
' Christian Langrock
' christian.langrock@actemium.de
 
Option Explicit
 
 Public Kartentyp As String
 Public Stationsnummer As Integer
Private pKWSBMK As String
Private pSignal As Integer
'Private pKartentyp As String
'Private pStationsnummer As Integer
Private pSteckplatz As Integer
Private pKanal As Integer
Private pAdress As String
Private pSegmentvorlage As String

''''''''''''''''''''''
' KWSBMK property
''''''''''''''''''''''
Public Property Get KWSBMK() As String
    KWSBMK = pKWSBMK
End Property

Public Property Let KWSBMK(ByVal Value As String)
    pKWSBMK = Value
End Property


''''''''''''''''''''''
' Signal property
''''''''''''''''''''''
Public Property Get Signal() As Integer
    Signal = pSignal
End Property

Public Property Let Signal(Value As Integer)
    pSignal = Value
End Property

''''''''''''''''''''''
' Stationsnummer property
''''''''''''''''''''''
'Public Property Get Stationsnummer() As Integer
 '   Stationsnummer = pStationsnummer
'End Property

'Public Property Let Stationsnummer(Value As Integer)
 '   pStationsnummer = Value
'End Property

''''''''''''''''''''''
' Kartentyp property
''''''''''''''''''''''
'Public Property Get Kartentyp() As String
'    Kartentyp = pKartentyp
'End Property

'Public Property Let Kartentyp(ByVal Value As String)
'    pKartentyp = Value
'End Property

''''''''''''''''''''''
' Steckplatz property
''''''''''''''''''''''
Public Property Get Steckplatz() As Integer
    Steckplatz = pSteckplatz
End Property

Public Property Let Steckplatz(Value As Integer)
    pSteckplatz = Value
End Property

''''''''''''''''''''''
' Kanal property
''''''''''''''''''''''
Public Property Get Kanal() As Integer
    Kanal = pKanal
End Property

Public Property Let Kanal(Value As Integer)
    pKanal = Value
End Property

''''''''''''''''''''''
' Adress property
''''''''''''''''''''''
Public Property Get Adress() As String
    Adress = pAdress
End Property

Public Property Let Adress(Value As String)
    pAdress = Value
End Property

''''''''''''''''''''''
' Segmentvorlage property
''''''''''''''''''''''
Public Property Get Segmentvorlage() As String
    Segmentvorlage = pSegmentvorlage
End Property

Public Property Let Segmentvorlage(Value As String)
    pSegmentvorlage = Value
End Property

Public Function CompareTo(per As cBelegung) As Long
    Dim i As Long

    If Me.Stationsnummer = per.Stationsnummer Then
        If Me.Kartentyp = per.Kartentyp Then
            i = 0
        ElseIf Me.Kartentyp < per.Kartentyp Then
            i = -1
        Else
            i = 1
        End If
    ElseIf Me.Stationsnummer < per.Stationsnummer Then
        i = -1
    Else
        i = 1
    End If

    CompareTo = i
End Function
