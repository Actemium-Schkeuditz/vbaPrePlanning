VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEinbauort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der Einbauorte
' V0.2
' 20.02.2020
' new property "BMK"
' Christian Langrock
' christian.langrock@actemium.de

'@folder(Kennzeichen.Einbauorte)
    
    Option Explicit

Private pStationsnummer As Integer
Private pEinbauort As String
Private pGeraetetyp As String
Private pBMK As String


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
' Stationsnummer property
''''''''''''''''''''''
Public Property Get Stationsnummer() As Integer
    Stationsnummer = pStationsnummer
End Property

Public Property Let Stationsnummer(ByVal Value As Integer)
    pStationsnummer = Value
End Property

''''''''''''''''''''''
' Geraetetyp property
''''''''''''''''''''''
Public Property Get Geraetetyp() As String
    Geraetetyp = pGeraetetyp
End Property

Public Property Let Geraetetyp(ByVal Value As String)
    pGeraetetyp = Value
End Property

''''''''''''''''''''''
' BMK property
''''''''''''''''''''''
Public Property Get BMK() As String
    BMK = pBMK
End Property

Public Property Let BMK(ByVal Value As String)
    pBMK = Value
End Property

