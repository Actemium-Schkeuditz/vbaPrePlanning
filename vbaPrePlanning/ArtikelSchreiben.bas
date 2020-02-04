Attribute VB_Name = "ArtikelSchreiben"
' Skript zum Übertragen der Artikeldaten
' V1.1
' 22.01.2020
' V1.1
' Änderung Zielzellen
' Christian Langrock
' christian.langrock@actemium.de

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
    Set ws1 = Worksheets(tabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        ' Spaltenbreiten anpassen
        ActiveSheet.Columns.Item("BK").Select
        Selection.ColumnWidth = 25


        '*********** Artikel schreiben und umkopieren ******************

        For i = 3 To zeilenanzahl
 
            If Cells(i, "F") <> "." Then
                'Prüfen ob am Ende ein Punkt ist
                If Right(Cells(i, "F"), 1) <> "." Then
                    'MsgBox ("kein Artikel: " + Right(Cells(i, "F"), 1))
                    'Else
                    Cells(i, "BK") = Cells(i, "F")
    
                    'Else
                    'MsgBox "kein Artikel"
                End If
            End If
            ' Artikel die nicht gewollt sind entfernen
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Siemens.7MH4138-6AA00-0BA0", vbNullString)
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Siemens.Siwarex WP321", "")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Siwarex WP321.7MH4138-6AA00-0BA0+BU15-P16+A0+2D", "")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Siemens.Sirius Act", "")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Stöbich.", vbNullString)
    
            ' ersetzen von falschen Ausdrücken
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Baumer", "BAU")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "ifm", "IFM")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Rechner Sensors", "RECH")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "MARTENS", "MAR")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Siemens", "SIE")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "Schmersal", "SCHM")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "IFM.IS 5001", "IFM.IS5001")
            Cells(i, "BK") = Replace(Cells(i, "BK"), "RECH.KA 0655", "RECH.KA0655")
    
        
        Next i

    End With
End Sub

