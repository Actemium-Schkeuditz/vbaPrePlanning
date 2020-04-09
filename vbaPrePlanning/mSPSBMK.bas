Attribute VB_Name = "mSPSBMK"
' Skript zur Ermittlung der SPS BMK´s
' Die Daten werden Kanalweise zugeordnet
' V0.5
' getestet
' 03.04.2020
' überflüssige Leerzeichen entfernt
'

' Christian Langrock
' christian.langrock@actemium.de

'@folder(Kennzeichen.SPS-BMK)
Option Explicit

Public Sub SPS_BMK()

    Dim ws1 As Worksheet
    Dim TabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim y As Long
    Dim spalteSPSKartentyp As Long
    Dim spalteSPSBMK As Long
    Dim spalteSPSSteckplatz As Long
    Dim ExcelConfig As New cExcelConfig
        
    ' Tabellen definieren
    TabelleDaten = ExcelConfig.TabelleDaten

    Set ws1 = Worksheets.[_Default](TabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
        'MsgBox zeilenanzahl

        For y = 1 To 6
            '*********** SPS-BMK erzeugen ******************
            ' Auswahl der Spalten pro SPS Kanal
            '**************** SPS-BMK für Signal 1 bis 6 ****************
            spalteSPSKartentyp = SpaltenBuchstaben2Int(ExcelConfig.Kartentyp) + 14 * (y - 1)
            spalteSPSBMK = SpaltenBuchstaben2Int(ExcelConfig.SPSBMK) + 14 * (y - 1)
            spalteSPSSteckplatz = SpaltenBuchstaben2Int(ExcelConfig.Steckplatz) + 14 * (y - 1)
          
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' Prüfen auf SPS-Typ
                ' ET200SP
                If Left$(.Cells.Item(i, spalteSPSKartentyp), 7) = "ET200SP" Then
                    '  MsgBox "Treffer" + Str(i)
                    ' erzeuge SPS-BMK wenn Steckplatz beschrieben
                    If .Cells.Item(i, spalteSPSSteckplatz) <> vbNullString Then
                        .Cells.Item(i, spalteSPSBMK) = Trim(str(.Cells.Item(i, spalteSPSSteckplatz) + 3)) + "K5"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        .Cells.Item(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' ET200AL
                ElseIf Left$(.Cells.Item(i, spalteSPSKartentyp), 7) = "ET200AL" Then
                    'MsgBox "Treffer" + str(i) + "nicht fertig programmiert"
                    If .Cells.Item(i, spalteSPSSteckplatz) <> vbNullString Then
                        'MsgBox "ET200AL" + str(i) + "nicht fertig programmiert, BMK prüfen"
                        .Cells.Item(i, spalteSPSBMK) = Trim(str(.Cells.Item(i, spalteSPSSteckplatz) + 3)) + "K5"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        .Cells.Item(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' CPX-  elektrisch
                ElseIf Left$(.Cells.Item(i, spalteSPSKartentyp), 4) = "CPX-" Then
                    'MsgBox "CPX " + Str(i) + "nicht fertig programmiert"
                    If .Cells.Item(i, spalteSPSSteckplatz) <> vbNullString Then
                        ' MsgBox "CPX" + Str(i) + "nicht fertig programmiert, BMK prüfen"
                        .Cells.Item(i, spalteSPSBMK) = Trim(str(.Cells.Item(i, spalteSPSSteckplatz) + 3)) + "KF2"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        .Cells.Item(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' IFM IO-LINK
                ElseIf .Cells.Item(i, spalteSPSKartentyp) = "IFM IO-LINK" Then
                    'MsgBox "IFM IO-LINK " + Str(i) + "nicht fertig programmiert"
                    If .Cells.Item(i, spalteSPSSteckplatz) <> vbNullString Then
                        ' MsgBox "CPX " + Str(i) + "nicht fertig programmiert, BMK prüfen"
                        .Cells.Item(i, spalteSPSBMK) = "1KF5" ' + Str(Cells(i, spalteSPSSteckplatz))
                    Else
                        ' makiere fehlende Steckplatz Daten
                        .Cells.Item(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                End If
            Next i
        Next y

    End With
End Sub


