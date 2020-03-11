VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnschluss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
' Klasse zum speichern der SPS Anschlussdaten
' V0.3
' 25.02.2020
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.SPS-Anschlüsse)
 
 Option Explicit
 

Public Kartentyp As cPLCdata
Private pKanal As Integer
Private pAnschluss1 As String
Private pAnschluss2 As String
Private pAnschluss3 As String
Private pAnschluss4 As String
Private pAnschlussM As String
Private pAnschlussVS As String

''''''''''''''''''''''
' Kartentyp property
''''''''''''''''''''''
 Private Sub Class_Initialize()
    Set Me.Kartentyp = New cPLCdata

End Sub

''''''''''''''''''''''
' Kanal property
''''''''''''''''''''''
Public Property Get Kanal() As Integer
    Kanal = pKanal
End Property

Public Property Let Kanal(ByVal Value As Integer)
    pKanal = Value
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
' Anschluss_M property
''''''''''''''''''''''
Public Property Get AnschlussM() As String
    AnschlussM = pAnschlussM
End Property

Public Property Let AnschlussM(ByVal Value As String)
    pAnschlussM = Value
End Property

''''''''''''''''''''''
' Anschluss_VS property
''''''''''''''''''''''
Public Property Get AnschlussVS() As String
    AnschlussVS = pAnschlussVS
End Property

Public Property Let AnschlussVS(ByVal Value As String)
    pAnschlussVS = Value
End Property


