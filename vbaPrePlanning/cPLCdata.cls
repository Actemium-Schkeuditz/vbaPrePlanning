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
' V0.3
' 20.02.2020
' new property
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.SPS-Konfig)
  
Option Explicit
Private pKartentyp As String
Private pPLCtyp As String
Private pSortingIdentifier As Long
Private pMaxChannelPerSlot As Long
Private pDatatyp As String
Private pSignalDirection As String
Private pInputAdressLength As Long               'per channel
Private pInputMinLength As Long                  'per slot
Private pInputMaxLength As Long                  'per slot
Private pInputAdressDiagnosticLength As Long     'per slot
Private pOutputAdressLength As Long              'per channel
Private pOutputMinLength As Long                 'per slot
Private pOutputMaxLength As Long                 'per slot
Private pOutputAdressDiagnosticLength As Long    'per slot
Private pVariableAdressRange As Boolean          'per slot
Private pSameAdressRangeForOutput As Boolean     'per slot

' Adresslänge pro Kanal in Einheit des Datendyps
' Diagnose pro Karte in Byte

Public Property Let Kartentyp(ByVal Value As String)
    pKartentyp = Value
    Select Case pKartentyp
    Case "ET200SP 4IO-LINK"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 1
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 1
        pInputMaxLength = 144
        pOutputMinLength = 0
        pOutputMaxLength = 128
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "ET200SP 8DI"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 2
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Input"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 1
        pInputMaxLength = 1
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 8DI HS"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 3
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Input"
        pInputAdressLength = 1                   'Bit
        pOutputAdressLength = 0
        pInputMinLength = 1
        pInputMaxLength = 1
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0         'Byte
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 8DI CNT"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 4
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
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
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 5
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4DO"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 6
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4RQ"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 7
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 2AI I"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 8
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 2AI U"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 9
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
        pInputAdressLength = 2
        pOutputAdressLength = 0
        pInputMinLength = 4
        pInputMaxLength = 4
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200SP 4AI PT100"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 10
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
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
    Case "ET200SP 2AO I"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 11
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 2
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 4
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "ET200SP 2AO U"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 12
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 2
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 4
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
    Case "ET200SP PTP"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 13
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
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
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 14
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Word"
        pSignalDirection = "InOut"
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
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 15
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Input"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 1
        pInputMaxLength = 1
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 5
        pOutputAdressDiagnosticLength = 4
        pVariableAdressRange = False
        pSameAdressRangeForOutput = True
    Case "ET200SP 4FDO"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 16
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 5
        pOutputAdressDiagnosticLength = 4
        pVariableAdressRange = False
        pSameAdressRangeForOutput = True
    Case "ET200SP STARTER"
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 17
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
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
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 18
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
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
        pPLCtyp = "ET200SP"
        pSortingIdentifier = 19
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
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
        pPLCtyp = "FESTO CPX"
        pSortingIdentifier = 29
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
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
        pPLCtyp = "FESTO CPX"
        pSortingIdentifier = 30
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Input"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 1
        pInputMaxLength = 1
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX-4DA"
        pPLCtyp = "FESTO CPX"
        pSortingIdentifier = 31
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "CPX-4AE-T"
        pPLCtyp = "FESTO CPX"
        pSortingIdentifier = 32
        pMaxChannelPerSlot = 3
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Input"
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
        pPLCtyp = "FESTO MPA"
        pSortingIdentifier = 43
        pMaxChannelPerSlot = 1
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
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
        pPLCtyp = "FESTO MPA"
        pSortingIdentifier = 41
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
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
        pPLCtyp = "FESTO MPA"
        pSortingIdentifier = 40
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
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
        pPLCtyp = "IFM IO-LINK MASTER"
        pSortingIdentifier = 50
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
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
        pPLCtyp = "IFM IO-LINK MASTER"
        pSortingIdentifier = 51
        pMaxChannelPerSlot = 4
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 4
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 1
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "AL1402"
        pPLCtyp = "IFM IO-LINK MASTER"
        pSortingIdentifier = 52
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
        pInputAdressLength = 32
        pOutputAdressLength = 32
        pInputMinLength = 4
        pInputMaxLength = 0
        pOutputMinLength = 4
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 1
        pOutputAdressDiagnosticLength = 1
        pVariableAdressRange = True
        pSameAdressRangeForOutput = True
    Case "ET200AL 16DI"
        pPLCtyp = "ET200AL"
        pSortingIdentifier = 60
        pMaxChannelPerSlot = 15
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Input"
        pInputAdressLength = 1
        pOutputAdressLength = 0
        pInputMinLength = 2
        pInputMaxLength = 2
        pOutputMinLength = 0
        pOutputMaxLength = 0
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "ET200AL 8DO"
        pPLCtyp = "ET200AL"
        pSortingIdentifier = 61
        pMaxChannelPerSlot = 7
        ' Adress
        pDatatyp = "Bit"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "FC302"
        pPLCtyp = "FU"
        pSortingIdentifier = 100
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
        pInputAdressLength = 30
        pOutputAdressLength = 30
        pInputMinLength = 30
        pInputMaxLength = 30
        pOutputMinLength = 30
        pOutputMaxLength = 30
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = True
    Case "COM800"
        pPLCtyp = "SIRIUS"
        pSortingIdentifier = 80
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "InOut"
        pInputAdressLength = 60     'pro Schalter 30Byte
        pOutputAdressLength = 4     'pro Schalter 2 Byte
        pInputMinLength = 30
        pInputMaxLength = 240
        pOutputMinLength = 2
        pOutputMaxLength = 16
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "FESTO MPA1S"
        pPLCtyp = "FESTO MPA"
        pSortingIdentifier = 150
        pMaxChannelPerSlot = 8
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
        pInputAdressLength = 0
        pOutputAdressLength = 1
        pInputMinLength = 0
        pInputMaxLength = 0
        pOutputMinLength = 1
        pOutputMaxLength = 1
        pInputAdressDiagnosticLength = 0
        pOutputAdressDiagnosticLength = 0
        pVariableAdressRange = False
        pSameAdressRangeForOutput = False
    Case "VMPA1-FB-SPU"
        pPLCtyp = "FESTO MPA"
        pSortingIdentifier = 151
        pMaxChannelPerSlot = 0
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = "Output"
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
    Case Else
        pPLCtyp = "unbestimmt"
        pSortingIdentifier = 200
        pMaxChannelPerSlot = 100
        ' Adress
        pDatatyp = "Byte"
        pSignalDirection = vbNullString
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
' Kartentyp
''''''''''''''''''''''
Public Property Get Kartentyp() As String
    Kartentyp = pKartentyp
End Property

''''''''''''''''''''''
' SPS-Typ
''''''''''''''''''''''
Public Property Get PLCTyp() As String
    PLCTyp = pPLCtyp
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

''''''''''''''''''''''
' SignalDirection
''''''''''''''''''''''
Public Property Get SignalDirection() As String
    SignalDirection = pSignalDirection
End Property

