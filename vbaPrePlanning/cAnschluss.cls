VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnschluss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
' Klasse zum speichern der SPS Anschlussdaten
' V0.2
' 28.01.2020
' Christian Langrock
' christian.langrock@actemium.de
 
 Option Explicit
 
'Public Kartentyp As String
'Public Kanal As Integer
Private pKartentyp As String
Private pKanal As Integer
Private pAnschluss_1 As String
Private pAnschluss_2 As String
Private pAnschluss_3 As String
Private pAnschluss_4 As String
Private pAnschluss_M As String
Private pAnschluss_VS As String

''''''''''''''''''''''
' Kartentyp property
''''''''''''''''''''''
Public Property Get Kartentyp() As String
    Kartentyp = pKartentyp
End Property

Public Property Let Kartentyp(ByVal Value As String)
    pKartentyp = Value
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
' Anschluss_1 property
''''''''''''''''''''''
Public Property Get Anschluss_1() As String
    Anschluss_1 = pAnschluss_1
End Property

Public Property Let Anschluss_1(Value As String)
    pAnschluss_1 = Value
End Property

''''''''''''''''''''''
' Anschluss_2 property
''''''''''''''''''''''
Public Property Get Anschluss_2() As String
    Anschluss_2 = pAnschluss_2
End Property

Public Property Let Anschluss_2(Value As String)
    pAnschluss_2 = Value
End Property

''''''''''''''''''''''
' Anschluss_3 property
''''''''''''''''''''''
Public Property Get Anschluss_3() As String
    Anschluss_3 = pAnschluss_3
End Property

Public Property Let Anschluss_3(Value As String)
    pAnschluss_3 = Value
End Property

''''''''''''''''''''''
' Anschluss_4 property
''''''''''''''''''''''
Public Property Get Anschluss_4() As String
    Anschluss_4 = pAnschluss_4
End Property

Public Property Let Anschluss_4(Value As String)
    pAnschluss_4 = Value
End Property

''''''''''''''''''''''
' Anschluss_M property
''''''''''''''''''''''
Public Property Get Anschluss_M() As String
    Anschluss_M = pAnschluss_M
End Property

Public Property Let Anschluss_M(Value As String)
    pAnschluss_M = Value
End Property

''''''''''''''''''''''
' Anschluss_VS property
''''''''''''''''''''''
Public Property Get Anschluss_VS() As String
    Anschluss_VS = pAnschluss_VS
End Property

Public Property Let Anschluss_VS(Value As String)
    pAnschluss_VS = Value
End Property

