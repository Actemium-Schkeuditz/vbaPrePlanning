Attribute VB_Name = "SeitenZahl"
Option Explicit
' Skript zum schreiben der Seitenzahlen
' V0.4
' 10.02.2020
' erste Funktion getestet, weitere Filter müssen noch rein
' Auslagern der Sortierfunktion
' Christian Langrock
' christian.langrock@actemium.de

' ToDO: testen
Public Sub SeitenZahlschreiben()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim Sortierspalte As String
    Dim Sortierspalte2 As String
    Dim SpalteSeitenzahl As String
    Dim SpalteAnlage As String
    Dim SpaltePneumatik As String
    Dim SpalteSegmentvorlage As String
    Dim KennzeichenOld As String
    Dim Seite As Integer
    Dim SeitePneumatik As Integer
    Dim answer As Integer
    
    On Error GoTo ErrorHandle
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets(tabelleDaten)
   
    'Application.ScreenUpdating = False


    Sortierspalte = "B"                          ' sortieren nach KWS-BMK
    Sortierspalte2 = "BQ"                        ' sortieren nach Einbauort
    SpalteSeitenzahl = "BR"
    SpalteAnlage = "C"
    SpaltePneumatik = "BB"
    SpalteSegmentvorlage = "BL"
    
    answer = MsgBox("Spalte BU Stationsnummern und Einbauorte schon geprüft?", vbQuestion + vbYesNo + vbDefaultButton2, "Prüfung der Daten")
    'Prüfe Stationsnummer
    If answer = vbYes Then
    
        ThisWorkbook.Worksheets(tabelleDaten).Activate
    
        ' Tabelle mit Daten bearbeiten
        With ws1
   
            ' Herausfinden der Anzahl der Zeilen
            zeilenanzahl = .Cells(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
            'MsgBox zeilenanzahl

            ' sortieren nach KWS-BMK
            ' nach Einbauort SPS Rack (Stationsnummer)
            ' Daten sortieren
            SortTable tabelleDaten, SpalteAnlage, Sortierspalte2, Sortierspalte
    

            ' ab hier Seitenzahlen vergeben
            ' sowie sortiert vergebeben, dabei berücksichtigen ob Pneumatik oder nicht
            KennzeichenOld = "Leerplatz"

            For i = 3 To zeilenanzahl
                If .Cells(i, SpalteAnlage) <> KennzeichenOld Then
                    SeitePneumatik = 1
                    Seite = 1
                End If
                KennzeichenOld = .Cells(i, SpalteAnlage)
                ' Prüfen ob nicht Segmentvorlage ohne Seite
                If .Cells(i, SpalteSegmentvorlage) <> "Sensor_ohne_SLP" Then
                    If .Cells(i, Sortierspalte2) <> vbNullString Then ' prüfen ob Sortierspalte nicht leer
                        If .Cells(i, SpaltePneumatik) = vbNullString Then
                            .Cells(i, SpalteSeitenzahl) = Seite
                            Seite = Seite + 1
                        ElseIf .Cells(i, SpaltePneumatik) <> vbNullString Then
                            .Cells(i, SpalteSeitenzahl) = SeitePneumatik
                            SeitePneumatik = SeitePneumatik + 1
                        Else
                            .Cells(i, SpalteSeitenzahl) = vbNullString
                        End If
                    Else
                        .Cells(i, SpalteSeitenzahl) = vbNullString
                    End If
                Else
                    .Cells(i, SpalteSeitenzahl) = vbNullString
                End If
            Next i
        End With
    End If

BeforeExit:
    ' Set rCell = Nothing
    'Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Fehler im Modul Sortieren.", vbCritical, "Error"
    Resume BeforeExit
End Sub

