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
' V0.2
' 11.02.2020
' Key ist neu
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Kanalbelegung)
 
Option Explicit
 
Private pKey As String
Private pKWSBMK As String
Private pSignal As Integer
Private pKartentyp As String
Private pStationsnummer As Integer
Private pSteckplatz As Integer
Private pKanal As Integer
Private pAdress As String
Private pSegmentvorlage As String
Private pSortingIdentifier As Long
Private pMaxChannelPerSlot As Long

''''''''''''''''''''''
' KWSBMK property
''''''''''''''''''''''
Public Property Get key() As String
    key = pKey
End Property

Public Property Let key(ByVal Value As String)
    pKey = Value
End Property


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

Public Property Let Signal(ByRef Value As Integer)
    pSignal = Value
End Property

''''''''''''''''''''''
' Stationsnummer property
''''''''''''''''''''''
Public Property Get Stationsnummer() As Integer
    Stationsnummer = pStationsnummer
End Property

Public Property Let Stationsnummer(ByRef Value As Integer)
    pStationsnummer = Value
End Property

''''''''''''''''''''''
' Kartentyp property
''''''''''''''''''''''
Public Property Get Kartentyp() As String
    Kartentyp = pKartentyp
End Property

Public Property Let Kartentyp(ByVal Value As String)
    pKartentyp = Value
    Select Case pKartentyp
    Case "ET200SP 4IO-LINK"
        pSortingIdentifier = 1
        pMaxChannelPerSlot = 3
    Case "ET200SP 8DI"
        pSortingIdentifier = 2
        pMaxChannelPerSlot = 7
    Case "ET200SP 8DI HS"
        pSortingIdentifier = 3
        pMaxChannelPerSlot = 7
    Case "ET200SP 8DI CNT"
        pSortingIdentifier = 4
        pMaxChannelPerSlot = 3
    Case "ET200SP 8DO"
        pSortingIdentifier = 5
        pMaxChannelPerSlot = 7
    Case "ET200SP 4DO"
        pSortingIdentifier = 6
        pMaxChannelPerSlot = 3
    Case "ET200SP 4RQ"
        pSortingIdentifier = 7
        pMaxChannelPerSlot = 3
    Case "ET200SP 2AI I"
        pSortingIdentifier = 8
        pMaxChannelPerSlot = 1
    Case "ET200SP 2AI U"
        pSortingIdentifier = 9
        pMaxChannelPerSlot = 1
    Case "ET200SP 4AI PT100"
        pSortingIdentifier = 10
        pMaxChannelPerSlot = 3
    Case "ET200SP 2AO I"
        pSortingIdentifier = 11
        pMaxChannelPerSlot = 1
    Case "ET200SP 2AO U"
        pSortingIdentifier = 12
        pMaxChannelPerSlot = 1
    Case "ET200SP PTP"
        pSortingIdentifier = 13
        pMaxChannelPerSlot = 0
    Case "ET200SP SIWAREX"
        pSortingIdentifier = 14
        pMaxChannelPerSlot = 0
    Case "ET200SP 8FDI"
        pSortingIdentifier = 15
        pMaxChannelPerSlot = 7
    Case "ET200SP 4FDO"
        pSortingIdentifier = 16
        pMaxChannelPerSlot = 3
    Case "ET200SP STARTER"
        pSortingIdentifier = 17
        pMaxChannelPerSlot = 0
    Case "ET200SP WENDESTARTER"
        pSortingIdentifier = 18
        pMaxChannelPerSlot = 0
        ' FESTO
    Case "CPX-8DE-D"
        pSortingIdentifier = 30
        pMaxChannelPerSlot = 7
    Case "CPX-4DA"
        pSortingIdentifier = 31
        pMaxChannelPerSlot = 3
    Case "CPX-4AE-T"
        pSortingIdentifier = 32
        pMaxChannelPerSlot = 3
    Case "CPX 2x3/2 mono"
        pSortingIdentifier = 40
        pMaxChannelPerSlot = 1
    Case "CPX 5/2 mono"
        pSortingIdentifier = 41
        pMaxChannelPerSlot = 1
    Case "CPX 5/2 bistabil"
        pSortingIdentifier = 42
        pMaxChannelPerSlot = 0
        'IFM IO-Link Master
    Case "IFM IO-LINK"
        pSortingIdentifier = 50
        pMaxChannelPerSlot = 7
    Case "ET200AL 16DI"
        pSortingIdentifier = 60
        pMaxChannelPerSlot = 15
    Case "ET200AL 8DO"
        pSortingIdentifier = 61
        pMaxChannelPerSlot = 7
    Case Else
        pSortingIdentifier = 200
        pMaxChannelPerSlot = 100
    End Select
    
End Property

''''''''''''''''''''''
' Steckplatz property
''''''''''''''''''''''
Public Property Get Steckplatz() As Integer
    Steckplatz = pSteckplatz
End Property

Public Property Let Steckplatz(ByRef Value As Integer)
    pSteckplatz = Value
End Property

''''''''''''''''''''''
' Kanal property
''''''''''''''''''''''
Public Property Get Kanal() As Integer
    Kanal = pKanal
End Property

Public Property Let Kanal(ByRef Value As Integer)
    pKanal = Value
End Property

''''''''''''''''''''''
' Adress property
''''''''''''''''''''''
Public Property Get Adress() As String
    Adress = pAdress
End Property

Public Property Let Adress(ByVal Value As String)
    pAdress = Value
End Property

''''''''''''''''''''''
' Segmentvorlage property
''''''''''''''''''''''
Public Property Get Segmentvorlage() As String
    Segmentvorlage = pSegmentvorlage
End Property

Public Property Let Segmentvorlage(ByRef Value As String)
    pSegmentvorlage = Value
End Property

''''''''''''''''''''''
' SortingIdentifier property
''''''''''''''''''''''
Public Property Get SortingIdentifier() As Long
    SortingIdentifier = pSortingIdentifier
End Property

''''''''''''''''''''''
' MaxChannelPerSlot property
''''''''''''''''''''''
Public Property Get MaxChannelPerSlot() As Long
    MaxChannelPerSlot = pMaxChannelPerSlot
End Property


Public Function CompareTo(ByRef per As cBelegung) As Long
    Dim i As Long
    'Sort first:    Stationsnummer
    'Sort second:   SortingIdentifier
    'Sort third:    KWSBMK
    If Me.Stationsnummer = per.Stationsnummer Then
        If Me.SortingIdentifier = per.SortingIdentifier Then
            If Me.KWSBMK = per.KWSBMK Then
                i = 0
            ElseIf Me.KWSBMK < per.KWSBMK Then
                i = -1
            Else
                i = 1
            End If
        ElseIf Me.SortingIdentifier < per.SortingIdentifier Then
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

Public Function CompareToSegment(ByRef per As cBelegung) As Long
    Dim i As Long
  
    'Sort first:    Segmentvorlage
    'Sort second:   KWSBMK
    'Sort third:    Signal
    If Me.Segmentvorlage = per.Segmentvorlage Then
        If Me.KWSBMK = per.KWSBMK Then
            If Me.Signal = per.Signal Then
                i = 0
            ElseIf Me.Signal < per.Signal Then
                i = -1
            Else
                i = 1
            End If
        ElseIf Me.KWSBMK < per.KWSBMK Then
            i = -1
        Else
            i = 1
        End If
    ElseIf Me.Segmentvorlage < per.Segmentvorlage Then
        i = -1
    Else
        i = 1
    End If

    CompareToSegment = i
End Function

