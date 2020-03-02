Attribute VB_Name = "mArtikelSchreiben"
' Skript zum Übertragen der Artikeldaten
' V1.1
' 22.01.2020
' Änderung Zielzellen
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.Artikel)
Option Explicit

Public Sub ArtikelBearbeiten()


    '@Ignore VariableNotUsed
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        ' Spaltenbreiten anpassen
        ActiveSheet.Columns.Item("BK").Select
        Selection.ColumnWidth = 25


        '*********** Artikel schreiben und umkopieren ******************

        For i = 3 To zeilenanzahl
 
            If .Cells.Item(i, "F") <> "." Then
                'Prüfen ob am Ende ein Punkt ist
                If Right$(.Cells.Item(i, "F"), 1) <> "." Then
                    'MsgBox ("kein Artikel: " + Right(Cells(i, "F"), 1))
                    'Else
                    .Cells.Item(i, "BK") = .Cells.Item(i, "F")
    
                    'Else
                    'MsgBox "kein Artikel"
                End If
            End If
            ' Artikel die nicht gewollt sind entfernen
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Siemens.7MH4138-6AA00-0BA0", vbNullString)
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Siemens.Siwarex WP321", vbNullString)
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Siwarex WP321.7MH4138-6AA00-0BA0+BU15-P16+A0+2D", vbNullString)
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Siemens.Sirius Act", vbNullString)
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Stöbich.", vbNullString)
    
            ' ersetzen von falschen Ausdrücken
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Baumer", "BAU")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "ifm", "IFM")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Rechner Sensors", "RECH")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "MARTENS", "MAR")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Siemens", "SIE")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "Schmersal", "SCHM")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "IFM.IS 5001", "IFM.IS5001")
            .Cells.Item(i, "BK") = Replace(.Cells.Item(i, "BK"), "RECH.KA 0655", "RECH.KA0655")
    
        
        Next i

    End With
End Sub

