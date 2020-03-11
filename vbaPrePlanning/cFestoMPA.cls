VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFestoMPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der Festo daten
' V0.1
' 11.03.2020
' neu
' Christian Langrock
' christian.langrock@actemium.de

'@folder(Daten.Festo)

Option Explicit

Private pAnschlussplatteMPA As Long
Private pZwischeneinspeisungMPA As Long
Private pSteckplatzMPA As Long
Private pKanalMPA As Long

''''''''''''''''''''''''''''''
' AnschlussplatteMPA property
'''''''''''''''''''''''''''''
Public Property Get AnschlussplatteMPA() As Long
    AnschlussplatteMPA = pAnschlussplatteMPA
End Property

Public Property Let AnschlussplatteMPA(ByVal Value As Long)
    pAnschlussplatteMPA = Value
End Property

''''''''''''''''''''''''''''''
' ZwischeneinspeisungMPA property
'''''''''''''''''''''''''''''
Public Property Get ZwischeneinspeisungMPA() As Long
    ZwischeneinspeisungMPA = pZwischeneinspeisungMPA
End Property

Public Property Let ZwischeneinspeisungMPA(ByVal Value As Long)
    pZwischeneinspeisungMPA = Value
End Property

''''''''''''''''''''''''''''''
' SteckplatzMPA property
'''''''''''''''''''''''''''''
Public Property Get SteckplatzMPA() As Long
    SteckplatzMPA = pSteckplatzMPA
End Property

Public Property Let SteckplatzMPA(ByVal Value As Long)
    pSteckplatzMPA = Value
End Property

''''''''''''''''''''''''''''''
' KanalMPA property
'''''''''''''''''''''''''''''
Public Property Get KanalMPA() As Long
    KanalMPA = pKanalMPA
End Property

Public Property Let KanalMPA(ByVal Value As Long)
    pKanalMPA = Value
End Property

Public Sub reset()
    pAnschlussplatteMPA = 0
    pZwischeneinspeisungMPA = 0
    pSteckplatzMPA = 0
    pKanalMPA = 0

End Sub

