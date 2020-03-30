Attribute VB_Name = "mArtikelSchreiben"
' Skript zum Übertragen der Artikeldaten
' V1.2
' 30.03.2020
' Änderung auf ExcelConfig
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.Artikel)
Option Explicit

Public Sub ArtikelBearbeiten()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim ExcelConfig As New cExcelConfig
      
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
        ActiveSheet.Columns.Item(ExcelConfig.Artikel).Select
        Selection.ColumnWidth = 25


        '*********** Artikel schreiben und umkopieren ******************

        For i = 3 To zeilenanzahl
 
            If .Cells.Item(i, ExcelConfig.ArtikelKWS) <> "." Then
                'Prüfen ob am Ende ein Punkt ist
                If Right$(.Cells.Item(i, ExcelConfig.ArtikelKWS), 1) <> "." Then
                    .Cells.Item(i, ExcelConfig.Artikel) = .Cells.Item(i, ExcelConfig.ArtikelKWS)
    
                End If
            End If
            ' Artikel die nicht gewollt sind entfernen
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Siemens.7MH4138-6AA00-0BA0", vbNullString)
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Siemens.Siwarex WP321", vbNullString)
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Siwarex WP321.7MH4138-6AA00-0BA0+BU15-P16+A0+2D", vbNullString)
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Siemens.Sirius Act", vbNullString)
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Stöbich.", vbNullString)
    
            ' ersetzen von falschen Ausdrücken
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Baumer", "BAU")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "ifm", "IFM")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Rechner Sensors", "RECH")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "MARTENS", "MAR")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Siemens", "SIE")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "Schmersal", "SCHM")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "IFM.IS 5001", "IFM.IS5001")
            .Cells.Item(i, ExcelConfig.Artikel) = Replace(.Cells.Item(i, ExcelConfig.Artikel), "RECH.KA 0655", "RECH.KA0655")
    
        
        Next i

    End With
End Sub

