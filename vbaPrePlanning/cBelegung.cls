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
' V0.4
' 12.03.2020
' new propertys
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Kanalbelegung)
 
Option Explicit
 
Private pKey As String
Private pKWSBMK As String
Private pAnlage As String
Private pEinbauort As String
Private pSeite As Long
Private pSignal As Integer
Public Kartentyp As cPLCdata

Private pStationsnummer As Integer
Private pSteckplatz As Integer
Private pSubSlot As Integer
Private pKanal As Integer
Private pAdress As String                        ' Adresse für Stromlaufplan
Private pSegmentvorlage As String
Private pInputAdressSize As Long
Private pOutputAdressSize As Long
Private pInputAdressPerSlot As Long
Private pOutputAdressPerSlot As Long
Private pAnschluss1 As String
Private pAnschluss2 As String
Private pAnschluss3 As String
Private pAnschluss4 As String
Private pAnschlussM As String
Private pAnschlussVS As String
Private pSPSBMK As String

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
' KWSBMK property
''''''''''''''''''''''
Public Property Get KWSBMK() As String
    KWSBMK = pKWSBMK
End Property

Public Property Let KWSBMK(ByVal Value As String)
    pKWSBMK = Value
End Property

''''''''''''''''''''''
' Anlage property
''''''''''''''''''''''
Public Property Get Anlage() As String
    Anlage = pAnlage
End Property

Public Property Let Anlage(ByVal Value As String)
    pAnlage = Value
End Property

''''''''''''''''''''''
' Einbauort property
''''''''''''''''''''''
Public Property Get Einbauort() As String
    Einbauort = pEinbauort
End Property

Public Property Let Einbauort(ByVal Value As String)
    pEinbauort = Value
End Property

''''''''''''''''''''''
' Seite property
''''''''''''''''''''''
Public Property Get Seite() As Long
    Seite = pSeite
End Property

Public Property Let Seite(ByRef Value As Long)
    pSeite = Value
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
' InputAdressSize
''''''''''''''''''''''
Public Property Get InputAdressSize() As Long
    InputAdressSize = pInputAdressSize
End Property

Public Property Let InputAdressSize(ByRef Value As Long)
    pInputAdressSize = Value
End Property

''''''''''''''''''''''
' OutputAdressSize
''''''''''''''''''''''
Public Property Get OutputAdressSize() As Long
    OutputAdressSize = pOutputAdressSize
End Property

Public Property Let OutputAdressSize(ByRef Value As Long)
    pOutputAdressSize = Value
End Property

''''''''''''''''''''''
' InputAdressPerSlot
''''''''''''''''''''''
Public Property Get InputAdressPerSlot() As Long
    InputAdressPerSlot = pInputAdressPerSlot
End Property

Public Property Let InputAdressPerSlot(ByRef Value As Long)
    pInputAdressPerSlot = Value
End Property

''''''''''''''''''''''
' OutputAdressPerSlot
''''''''''''''''''''''
Public Property Get OutputAdressPerSlot() As Long
    OutputAdressPerSlot = pOutputAdressPerSlot
End Property

Public Property Let OutputAdressPerSlot(ByRef Value As Long)
    pOutputAdressPerSlot = Value
End Property

''''''''''''''''''''''
' Anschluss_1 property
''''''''''''''''''''''
Public Property Get Anschluss1() As String
    Anschluss1 = pAnschluss1
End Property

Public Property Let Anschluss1(ByVal Value As String)
    pAnschluss1 = Value
End Property

''''''''''''''''''''''
' Anschluss_2 property
''''''''''''''''''''''
Public Property Get Anschluss2() As String
    Anschluss2 = pAnschluss2
End Property

Public Property Let Anschluss2(ByVal Value As String)
    pAnschluss2 = Value
End Property

''''''''''''''''''''''
' Anschluss_3 property
''''''''''''''''''''''
Public Property Get Anschluss3() As String
    Anschluss3 = pAnschluss3
End Property

Public Property Let Anschluss3(ByVal Value As String)
    pAnschluss3 = Value
End Property

''''''''''''''''''''''
' Anschluss_4 property
''''''''''''''''''''''
Public Property Get Anschluss4() As String
    Anschluss4 = pAnschluss4
End Property

Public Property Let Anschluss4(ByVal Value As String)
    pAnschluss4 = Value
End Property

''''''''''''''''''''''
' AnschlussVS property
''''''''''''''''''''''
Public Property Get AnschlussVS() As String
    AnschlussVS = pAnschlussVS
End Property

Public Property Let AnschlussVS(ByVal Value As String)
    pAnschlussVS = Value
End Property

''''''''''''''''''''''
' AnschlussM property
''''''''''''''''''''''
Public Property Get AnschlussM() As String
    AnschlussM = pAnschlussM
End Property

Public Property Let AnschlussM(ByVal Value As String)
    pAnschlussM = Value
End Property

''''''''''''''''''''''
' SPSBMK
''''''''''''''''''''''
Public Property Get SPSBMK() As String
    SPSBMK = pSPSBMK
End Property

Public Property Let SPSBMK(ByVal Value As String)
    pSPSBMK = Value
End Property

Public Function CompareTo(ByRef per As cBelegung) As Long
    Dim i As Long
    'Sort first:    Stationsnummer
    'Sort second:   SortingIdentifier
    'Sort third:    KWSBMK
    If Me.Stationsnummer = per.Stationsnummer Then
        'If Me.SortingIdentifier = per.SortingIdentifier Then
        If Me.Kartentyp.SortingIdentifier = per.Kartentyp.SortingIdentifier Then
            If Me.KWSBMK = per.KWSBMK Then
                i = 0
            ElseIf Me.KWSBMK < per.KWSBMK Then
                i = -1
            Else
                i = 1
            End If
            'elseif Me.SortingIdentifier < per.SortingIdentifier Then
        ElseIf Me.Kartentyp.SortingIdentifier < per.Kartentyp.SortingIdentifier Then
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

Public Function CompareToChannel(ByRef per As cBelegung) As Long
    Dim i As Long
  
    'Sort first:    Station
    'Sort second:   Slot
    'Sort third:    Channel
    If Me.Stationsnummer = per.Stationsnummer Then
        If Me.Steckplatz = per.Steckplatz Then
            If Me.Kanal = per.Kanal Then
                i = 0
            ElseIf Me.Kanal < per.Kanal Then
                i = -1
            Else
                i = 1
            End If
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

    CompareToChannel = i
End Function

Public Function CompareForPagenumbers(ByRef per As cBelegung) As Long
    Dim i As Long
  
    'Sort first:    higher level
    'Sort second:   monting location
    'Sort third:    KWSBMK
    If Me.Anlage = per.Anlage Then
        If Me.Einbauort = per.Einbauort Then
            If Me.KWSBMK = per.KWSBMK Then
                i = 0
            ElseIf Me.KWSBMK < per.KWSBMK Then
                i = -1
            Else
                i = 1
            End If
        ElseIf Me.Einbauort < per.Einbauort Then
            i = -1
        Else
            i = 1
        End If
    ElseIf Me.Anlage < per.Anlage Then
        i = -1
    Else
        i = 1
    End If

    CompareForPagenumbers = i
End Function


