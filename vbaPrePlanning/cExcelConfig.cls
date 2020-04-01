VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExcelConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasse zum speichern der Excelkonfiguration
' V0.1
' 30.03.2020
' new class
' Christian Langrock
' christian.langrock@actemium.de

'@folder (Daten.Excel)
  
Option Explicit
Private pAnlage As String
Private pSympolischeAdresse As String
Private pArtikel As String
Private pArtikelKWS As String
Private pSegmentdefinition As String
Private pStationsnummerKWS As String
Private pSegmentvorlage As String
Private pEingangsAdress As String
Private pAusgangsAdress As String
Private pStationsnummer As String
Private pEinbauortEinzel As String
Private pEinbauort As String
Private pSeite As String
Private pSPSRackEinbauort As String
Private pSPSRackBMK As String

'Signale 1 bis 6
Private pStationsnummerSignal_1 As String
Private pSPSRackBMKSignal_1 As String
Private pSPSBMK As String
Private pSteckplatz As String
Private pKanal As String
Private pKartentyp As String


Private Sub Class_Initialize()
    
' Spalten definieren
    pAnlage = "C"
    pArtikelKWS = "F"
    pStationsnummerKWS = "BC"
        pEingangsAdress = "BD"
    pAusgangsAdress = "BE"
    pSympolischeAdresse = "BJ"
    pArtikel = "BK"
    pSegmentvorlage = "BL"
    pSegmentdefinition = "BP"

    pEinbauortEinzel = "BR"
    pEinbauort = "BQ"
    pSeite = "BS"
    pStationsnummer = "BV"
    pSPSRackEinbauort = "BW"
    pSPSRackBMK = "BX"
  
    pStationsnummerSignal_1 = "BY"
    pSPSRackBMKSignal_1 = "BZ"
    pKartentyp = "CB"
    pSPSBMK = "CC"
    pSteckplatz = "CD"
    pKanal = "CE"
    
End Sub

''''''''''''''''''''''
' Anlage
''''''''''''''''''''''
Public Property Get Anlage() As String
    Anlage = pAnlage
End Property

''''''''''''''''''''''
' ArtikelKWS
''''''''''''''''''''''
Public Property Get ArtikelKWS() As String
    ArtikelKWS = pArtikelKWS
End Property

''''''''''''''''''''''
'StationsnummerKWS
''''''''''''''''''''''
Public Property Get StationsnummerKWS() As String
    StationsnummerKWS = pStationsnummerKWS
End Property

''''''''''''''''''''''
' Artikel
''''''''''''''''''''''
Public Property Get Artikel() As String
    Artikel = pArtikel
End Property

''''''''''''''''''''''
' SympolischeAdresse property
''''''''''''''''''''''
Public Property Get SympolischeAdresse() As String
    SympolischeAdresse = pSympolischeAdresse
End Property

''''''''''''''''''''''
' Segmentvorlage property
''''''''''''''''''''''
Public Property Get Segmentvorlage() As String
    Segmentvorlage = pSegmentvorlage
End Property

''''''''''''''''''''''
' Segmentdefinition property
''''''''''''''''''''''
Public Property Get Segmentdefinition() As String
    Segmentdefinition = pSegmentdefinition
End Property

''''''''''''''''''''''
' pEingangsAdress property
''''''''''''''''''''''
Public Property Get EingangsAdress() As String
    EingangsAdress = pEingangsAdress
End Property

''''''''''''''''''''''
' AusgangsAdress property
''''''''''''''''''''''
Public Property Get AusgangsAdress() As String
    AusgangsAdress = pAusgangsAdress
End Property

''''''''''''''''''''''
' Stationsnummer
''''''''''''''''''''''
Public Property Get Stationsnummer() As String
    Stationsnummer = pStationsnummer
End Property
''''''''''''''''''''''
'pSPSRackEinbauort
''''''''''''''''''''''
Public Property Get SPSRackEinbauort() As String
    SPSRackEinbauort = pSPSRackEinbauort
End Property

''''''''''''''''''''''
'pSPSRackBMK
''''''''''''''''''''''
Public Property Get SPSRackAnlage() As String
    SPSRackAnlage = pSPSRackBMK
End Property
''''''''''''''''''''''
' pEinbauortEinzel
''''''''''''''''''''''
Public Property Get EinbauortEinzel() As String
    EinbauortEinzel = pEinbauortEinzel
End Property

''''''''''''''''''''''
' pEinbauort
''''''''''''''''''''''
Public Property Get Einbauort() As String
    Einbauort = pEinbauort
End Property

''''''''''''''''''''''
' pSeite
''''''''''''''''''''''
Public Property Get Seite() As String
    Seite = pSeite
End Property

''''''''''''''''''''''
'StationsnummerSignal_1
''''''''''''''''''''''
Public Property Get StationsnummerSignal_1() As String
    StationsnummerSignal_1 = pStationsnummerSignal_1
End Property

''''''''''''''''''''''
'pSPSRackBMKSignal_1
''''''''''''''''''''''
Public Property Get SPSRackBMKSignal_1() As String
    SPSRackBMKSignal_1 = pSPSRackBMKSignal_1
End Property
''''''''''''''''''''''
'SPSBMK property
''''''''''''''''''''''
Public Property Get SPSBMK() As String
    SPSBMK = pSPSBMK
End Property

''''''''''''''''''''''
'Steckplatz property
''''''''''''''''''''''
Public Property Get Steckplatz() As String
    Steckplatz = pSteckplatz
End Property

''''''''''''''''''''''
' Kanal property
''''''''''''''''''''''
Public Property Get Kanal() As String
    Kanal = pKanal
End Property

''''''''''''''''''''''
'Kartentyp property
''''''''''''''''''''''
Public Property Get Kartentyp() As String
    Kartentyp = pKartentyp
End Property

''''''''''''''''''''''

