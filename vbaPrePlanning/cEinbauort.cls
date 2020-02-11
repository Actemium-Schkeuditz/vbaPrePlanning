VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEinbauort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Klasse zum speichern der Einbauorte
' V0.1
' 28.01.2020
' Christian Langrock
' christian.langrock@actemium.de

'@folder(Kennzeichen.Einbauorte)
    
Private pStationsnummer As Integer
Private pEinbauort As String

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

