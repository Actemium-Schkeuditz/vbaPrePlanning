VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPLCconfigData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der SPS Stationskonfiguratuion
' V0.2
' 17.02.2020
' new/rename Propertys
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Konfig)

 Option Explicit
 
 Private pKey As String     ' verwenden wenn aus Excelliste gelesen wird
 Public Kartentyp As cPLCdata
 Private pStationsnummer As Integer
 Private pSteckplatz As Integer
 Private pSubSlot As Integer
 Private pFirstInputAdress As Long
 Private pFirstOutputAdress As Long
 Private pReserveChannelsBefor As Long
 Private pReserveChannelsAfter As Long
 Private pReserveChannelperSlot As Long
 Private pReserveSlot As Long
 
 Private Sub Class_Initialize()
    Set Me.Kartentyp = New cPLCdata
End Sub

''''''''''''''''''''''
' Key property
''''''''''''''''''''''
Public Property Get Key() As String
    Key = pKey
End Property

Public Property Let Key(ByVal Value As String)
    pKey = Value
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
' Steckplatz property
''''''''''''''''''''''
Public Property Get Steckplatz() As Integer
    Steckplatz = pSteckplatz
End Property

Public Property Let Steckplatz(ByRef Value As Integer)
    pSteckplatz = Value
End Property

''''''''''''''''''''''
' SubSlot property
''''''''''''''''''''''
Public Property Get SubSlot() As Integer
    SubSlot = pSubSlot
End Property

Public Property Let SubSlot(ByRef Value As Integer)
    pSubSlot = Value
End Property

''''''''''''''''''''''
' pFirstInputAdress
''''''''''''''''''''''
Public Property Get FirstInputAdress() As Long
    FirstInputAdress = pFirstInputAdress
End Property

Public Property Let FirstInputAdress(ByRef Value As Long)
    pFirstInputAdress = Value
End Property

''''''''''''''''''''''
' pFirstOutputAdress
''''''''''''''''''''''
Public Property Get FirstOutputAdress() As Long
    FirstOutputAdress = pFirstOutputAdress
End Property

Public Property Let FirstOutputAdress(ByRef Value As Long)
    pFirstOutputAdress = Value
End Property

''''''''''''''''''''''
' OffsetSlot
''''''''''''''''''''''
Public Property Get ReserveChannelsBefor() As Long
    ReserveChannelsBefor = pReserveChannelsBefor
End Property

Public Property Let ReserveChannelsBefor(ByRef Value As Long)
    pReserveChannelsBefor = Value
End Property
''''''''''''''''''''''
' ReserveSlot
''''''''''''''''''''''
Public Property Get ReserveChannelsAfter() As Long
    ReserveChannelsAfter = pReserveChannelsAfter
End Property

Public Property Let ReserveChannelsAfter(ByRef Value As Long)
    pReserveChannelsAfter = Value
End Property
''''''''''''''''''''''
' ReserveChannel
''''''''''''''''''''''
Public Property Get ReserveChannelPerSlot() As Long
    ReserveChannelPerSlot = pReserveChannelperSlot
End Property

Public Property Let ReserveChannelPerSlot(ByRef Value As Long)
    pReserveChannelperSlot = Value
End Property

''''''''''''''''''''''
' ReserveSlot
''''''''''''''''''''''
Public Property Get ReserveSlot() As Long
    ReserveSlot = pReserveSlot
End Property

Public Property Let ReserveSlot(ByRef Value As Long)
    pReserveSlot = Value
End Property

'#########################################################################

Public Function ComparePLCconfigTo(ByRef per As cPLCconfigData) As Long
    Dim i As Long
    'Sort first:    Stationsnummer
    'Sort second:   Steckplatz

    If Me.Stationsnummer = per.Stationsnummer Then
        If Me.Steckplatz = per.Steckplatz Then
                i = 0
        ElseIf Me.Steckplatz < per.Steckplatz Then
            i = -1
        Else
            i = 1
        End If
    ElseIf Me.Stationsnummer < per.Stationsnummer Then
        i = -1
    Else
        i = 1
    End If

    ComparePLCconfigTo = i
End Function

