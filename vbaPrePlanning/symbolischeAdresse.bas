Attribute VB_Name = "symbolischeAdresse"
' Skript zur Ermittlung der symbolischen Adressen
' V1.2
'22.01.2020
'angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de

Option Explicit

Public Sub symbolische_Adresse()


    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
      
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

        '*********** symbolische Adresse erzeugen und umkopieren ******************


        ' Spaltenbreiten anpassen
        ActiveSheet.Columns.Item("BJ").Select
        Selection.ColumnWidth = 35


        For i = 3 To zeilenanzahl
            Cells(i, "BJ") = LTrim(Cells(i, "B")) 'führende Leerzeichen entfernen
        Next i

    End With
End Sub

