Attribute VB_Name = "SPSBMK"
' Skript zur Ermittlung der SPS BMK´s
' Die Daten werden Kanalweise zugeordnet
' V0.1
' nicht fertig
' 22.01.2020
' angepasst für MH04
'
' Christian Langrock
' christian.langrock@actemium.de

Option Explicit

Public Sub SPS_BMK()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim y As Integer
    Dim spalteSPSKartentyp As String
    Dim spalteSPSBMK As String
    Dim spalteSPSSteckplatz As String
    Dim spalteSPSKanal As String
      
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

        For y = 1 To 5
            '*********** SPS-BMK erzeugen ******************

            ' Auswahl der Spalten pro SPS Kanal
            If y = 1 Then
                '**************** SPS-BMK für Signal 1 ****************
                spalteSPSKartentyp = "BY"
                spalteSPSBMK = "BZ"
                spalteSPSSteckplatz = "CA"
                spalteSPSKanal = "CB"
            ElseIf y = 2 Then
                '**************** SPS-BMK für Signal 2 ****************
                spalteSPSKartentyp = "CK"
                spalteSPSBMK = "CL"
                spalteSPSSteckplatz = "CM"
                spalteSPSKanal = "CN"
            ElseIf y = 3 Then
                '**************** SPS-BMK für Signal 3 ****************
                spalteSPSKartentyp = "CW"
                spalteSPSBMK = "CX"
                spalteSPSSteckplatz = "CY"
                spalteSPSKanal = "CZ"
            ElseIf y = 4 Then
                '**************** SPS-BMK für Signal 4 ****************
                spalteSPSKartentyp = "DI"
                spalteSPSBMK = "DJ"
                spalteSPSSteckplatz = "DK"
                spalteSPSKanal = "DL"
            ElseIf y = 5 Then
                '**************** SPS-BMK für Signal 5 ****************
                spalteSPSKartentyp = "DU"
                spalteSPSBMK = "DV"
                spalteSPSSteckplatz = "DW"
                spalteSPSKanal = "DX"
                'MsgBox "kein Fehler"
            Else
                MsgBox "Fehler SPS-BMK erzeugen"
            End If
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' Prüfen auf SPS-Typ
                ' ET200SP
                If Left(Cells(i, spalteSPSKartentyp), 7) = "ET200SP" Then
                    '  MsgBox "Treffer" + Str(i)
                    ' erzeuge SPS-BMK wenn Steckplatz beschrieben ist
                    If Cells(i, spalteSPSSteckplatz) <> vbNullString Then
                        Cells(i, spalteSPSBMK) = str(Cells(i, spalteSPSSteckplatz) + 3) + "K5"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        Cells(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' ET200AL
                ElseIf Left(Cells(i, spalteSPSKartentyp), 7) = "ET200AL" Then
                    MsgBox "Treffer" + str(i) + "nicht fertig programmiert"
                    If Cells(i, spalteSPSSteckplatz) <> vbNullString Then
                        MsgBox "ET200AL" + str(i) + "nicht fertig programmiert, BMK prüfen"
                        Cells(i, spalteSPSBMK) = str(Cells(i, spalteSPSSteckplatz) + 3) + "K5"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        Cells(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' CPX-  elektrisch
                ElseIf Left(Cells(i, spalteSPSKartentyp), 4) = "CPX-" Then
                    'MsgBox "CPX " + Str(i) + "nicht fertig programmiert"
                    If Cells(i, spalteSPSSteckplatz) <> vbNullString Then
                        ' MsgBox "CPX" + Str(i) + "nicht fertig programmiert, BMK prüfen"
                        Cells(i, spalteSPSBMK) = str(Cells(i, spalteSPSSteckplatz) + 3) + "KF2"
                    Else
                        ' makiere fehlende Steckplatz Daten
                        Cells(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' "CPX "  pneumatisch
                ElseIf Left(Cells(i, spalteSPSKartentyp), 4) = "CPX " Then
                    'MsgBox "CPX " + Str(i) + "nicht fertig programmiert"
                    If Cells(i, spalteSPSSteckplatz) <> vbNullString Then
                        ' MsgBox "CPX " + Str(i) + "nicht fertig programmiert, BMK prüfen"
                        Cells(i, spalteSPSBMK) = "KH" + str(Cells(i, spalteSPSSteckplatz))
                    Else
                        ' makiere fehlende Steckplatz Daten
                        Cells(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                    ' IFM IO-LINK
                ElseIf Cells(i, spalteSPSKartentyp) = "IFM IO-LINK" Then
                    'MsgBox "IFM IO-LINK " + Str(i) + "nicht fertig programmiert"
                    If Cells(i, spalteSPSSteckplatz) <> vbNullString Then
                        ' MsgBox "CPX " + Str(i) + "nicht fertig programmiert, BMK prüfen"
                        Cells(i, spalteSPSBMK) = "1KF5" ' + Str(Cells(i, spalteSPSSteckplatz))
                    Else
                        ' makiere fehlende Steckplatz Daten
                        Cells(i, spalteSPSSteckplatz).Interior.ColorIndex = 3
                    End If
                End If
            Next i
        Next y




 


 

    End With
End Sub

