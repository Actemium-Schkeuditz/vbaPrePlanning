Attribute VB_Name = "SPSBelegung"
Option Explicit
' Skript zum zuweisen der Kanäle
' V0.1
' 10.02.2020
' erste Funktion getestet, weitere Filter müssen noch rein
' Christian Langrock
' christian.langrock@actemium.de

' ToDO: bisher nur Signal 1, Weitere Funktionen einfügen und testen
Public Sub SPSBelegung()

 Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim Sortierspalte As String
    Dim Sortierspalte2 As String
    
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets(tabelleDaten)
   
    Application.ScreenUpdating = False

    
     Sortierspalte = "BU"                          ' sortieren nach Stationsnummer
    Sortierspalte2 = "BY"                        ' sortieren nach Einbauort
    
    '### Sortieren der Daten nach Stationsnummer und Kartentyp ####
    SortTable tabelleDaten, Sortierspalte, Sortierspalte2

End Sub
