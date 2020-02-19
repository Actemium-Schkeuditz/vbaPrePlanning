VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPLCdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der SPS Kartendaten
' V0.2
' 19.02.2020
' new property
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Konfig)
  
  Option Explicit
Private pKartentyp As String
Private pFirstInputAdress As Long
Private pFirstOutputAdress As Long
Private pSortingIdentifier As Long
Private pMaxChannelPerSlot As Long
Private pDatatyp As String
Private pInputAdressLength As Long                  'per Channel
Private pInputMinLength As Long                     'per Channel
Private pInputMaxLength As Long                     'per Channel
Private pInputAdressDiagnosticLength As Long        'per Slot
Private pOutputAdressLength As Long                 'per Channel
Private pOutputMinLength As Long                    'per Channel
Private pOutputMaxLength As Long                    'per Channel
Private pOutputAdressDiagnosticLength As Long       'per Slot
Private pVariableAdressRange As Boolean             'per Slot
Private pSameAdressRangeForOutput As Boolean                  'per Slot

' Adresslänge pro Kanal in Einheit des Datendyps
' Diagnose pro Karte in Byte


Public Property Let Kartentyp(ByVal Value As String)
    pKartentyp = Value
    Select Case pKartentyp
    Case "ET200SP 4IO-LINK"
        
        pSortingIdentifier = 1
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 8
        pInputMaxLength = 64
        pOutputMinLength = 8
        pOutputMaxLength = 64
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "ET200SP 8DI"
        pSortingIdentifier = 2
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 8DI HS"
        pSortingIdentifier = 3
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1                   'Bit
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1         'Byte
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 8DI CNT"
        pSortingIdentifier = 4
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 6
        pOutputAdressLength = 5
        pInputMinLength = 25
        pInputMaxLength = 25
        pOutputMinLength = 20
        pOutputMaxLength = 20
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 8DO"
        pSortingIdentifier = 5
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 8
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4DO"
        pSortingIdentifier = 6
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 8
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4RQ"
        pSortingIdentifier = 7
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 8
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 2AI I"
        pSortingIdentifier = 8
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 2AI U"
        pSortingIdentifier = 9
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4AI PT100"
        pSortingIdentifier = 10
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 2AO I"
        pSortingIdentifier = 11
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 0
        pOutputAdressLength = 2
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 4
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "ET200SP 2AO U"
        pSortingIdentifier = 12
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 0
        pOutputAdressLength = 2
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 4
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "ET200SP PTP"
        pSortingIdentifier = 13
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 8
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP SIWAREX"
        pSortingIdentifier = 14
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 16
        pOutputAdressLength = 16
        pInputMinLength = 16
        pInputMaxLength = 16
        pOutputMinLength = 16
        pOutputMaxLength = 16
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = True
    Case "ET200SP 8FDI"
        pSortingIdentifier = 15
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 5
        pOutputAdressDiagnosticLength = 4
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4FDO"
        pSortingIdentifier = 16
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 8
        pInputAdressDiagnosticLength = 4
        pOutputAdressDiagnosticLength = 5
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP STARTER"
        pSortingIdentifier = 17
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 4
        pOutputAdressLength = 2
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 2
        pOutputMaxLength = 2
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP WENDESTARTER"
        pSortingIdentifier = 18
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 4
        pOutputAdressLength = 2
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 2
        pOutputMaxLength = 2
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP RESERVE"
        pSortingIdentifier = 19
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 0
        pOutputAdressLength = 0
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
        ' FESTO
    Case "FB43"
        pSortingIdentifier = 29
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 0
        pOutputAdressLength = 0
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 2
        pOutputAdressDiagnosticLength = 2
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX-8DE-D"
        pSortingIdentifier = 30
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX-4DA"
        pSortingIdentifier = 31
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 8
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX-4AE-T"
        pSortingIdentifier = 32
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 8
        pInputMaxLength = 8
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX 2x3/2 mono"
        pSortingIdentifier = 43
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 2
        pOutputMaxLength = 2
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "CPX 5/2 mono"
        pSortingIdentifier = 41
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 2
        pOutputMaxLength = 2
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "CPX 5/2 bistabil"
        pSortingIdentifier = 40
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 2
        pOutputMaxLength = 2
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
        'IFM IO-Link Master
    Case "IFM IO-LINK"
        pSortingIdentifier = 50
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 4
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "AL1400"
        pSortingIdentifier = 51
        pMaxChannelPerSlot = 4
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 4
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "AL1402"
        pSortingIdentifier = 52
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 4
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "ET200AL 16DI"
        pSortingIdentifier = 60
        pMaxChannelPerSlot = 15
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 16
        pInputMaxLength = 0
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 2
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200AL 8DO"
        pSortingIdentifier = 61
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 8
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case Else
        pSortingIdentifier = 200
        pMaxChannelPerSlot = 100
        ' Adress
        pDatatyp = "Byte"
        pInputAdressLength = 100
        pOutputAdressLength = 100
        pInputMinLength = 244
        pInputMaxLength = 0
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 2
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    End Select
    
End Property

''''''''''''''''''''''
' FirstInputAdress
''''''''''''''''''''''
Public Property Get FirstInputAdress() As Long
    FirstInputAdress = pFirstInputAdress
End Property

Public Property Let FirstInputAdress(ByVal Value As Long)
    pFirstInputAdress = Value
End Property
''''''''''''''''''''''
' FirstOutputAdress
''''''''''''''''''''''
Public Property Get FirstOutputAdress() As Long
    FirstOutputAdress = pFirstOutputAdress
End Property

Public Property Let FirstOutputAdress(ByVal Value As Long)
    pFirstOutputAdress = Value
End Property


''''''''''''''''''''''
' Kartentyp
''''''''''''''''''''''
Public Property Get Kartentyp() As String
Kartentyp = pKartentyp
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

''''''''''''''''''''''
' Datatyp
''''''''''''''''''''''
Public Property Get Datatyp() As String
    Datatyp = pDatatyp
End Property

''''''''''''''''''''''
' InputAdressLength
''''''''''''''''''''''
Public Property Get InputAdressLength() As Long
    InputAdressLength = pInputAdressLength
End Property

''''''''''''''''''''''
' OutputAdressLength
''''''''''''''''''''''
Public Property Get OutputAdressLength() As Long
    OutputAdressLength = pOutputAdressLength
End Property

''''''''''''''''''''''
' InputAdressDiagnosticLength
''''''''''''''''''''''
Public Property Get InputAdressDiagnosticLength() As Long
    InputAdressDiagnosticLength = pInputAdressDiagnosticLength
End Property

''''''''''''''''''''''
' OutputAdressDiagnosticLength
''''''''''''''''''''''
Public Property Get OutputAdressDiagnosticLength() As Long
    OutputAdressDiagnosticLength = pOutputAdressDiagnosticLength
End Property

''''''''''''''''''''''
' FixedAdressRangeForOutput
''''''''''''''''''''''
Public Property Get VariableAdressRange() As Boolean
    VariableAdressRange = pVariableAdressRange
End Property
''''''''''''''''''''''
' InputMinLength
''''''''''''''''''''''
Public Property Get InputMinLength() As Long
    InputMinLength = pInputMinLength
End Property

''''''''''''''''''''''
' InputMaxLength
''''''''''''''''''''''
Public Property Get InputMaxLength() As Long
    InputMaxLength = pInputMaxLength
End Property

''''''''''''''''''''''
' OutputMinLength
''''''''''''''''''''''
Public Property Get OutputMinLength() As Long
    OutputMinLength = pOutputMinLength
End Property

''''''''''''''''''''''
' OutputMaxLength
''''''''''''''''''''''
Public Property Get OutputMaxLength() As Long
    OutputMaxLength = pOutputMaxLength
End Property

''''''''''''''''''''''
' SameAdressRangeForOutput
''''''''''''''''''''''
Public Property Get SameAdressRangeForOutput() As Boolean
    SameAdressRangeForOutput = pSameAdressRangeForOutput
End Property
