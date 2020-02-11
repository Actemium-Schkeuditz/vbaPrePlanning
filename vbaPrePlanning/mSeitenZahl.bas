Attribute VB_Name = "mSeitenZahl"
Option Explicit
' Skript zum schreiben der Seitenzahlen
' V0.4
' 10.02.2020
' erste Funktion getestet, weitere Filter müssen noch rein
' Auslagern der Sortierfunktion
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.Seitenzahl)

' ToDO: testen
Public Sub SeitenZahlschreiben()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim Sortierspalte As String
    Dim Sortierspalte2 As String
    Dim SpalteSeitenzahl As String
    Dim SpalteAnlage As String
    Dim SpaltePneumatik As String
    Dim SpalteSegmentvorlage As String
    Dim KennzeichenOld As String
    Dim Seite As Long
    Dim SeitePneumatik As Long
    Dim answer As Long
    
    On Error GoTo ErrorHandle
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
   
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
    
        ThisWorkbook.Worksheets.[_Default](tabelleDaten).Activate
    
        ' Tabelle mit Daten bearbeiten
        With ws1
   
            ' Herausfinden der Anzahl der Zeilen
            zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gezählt
            'MsgBox zeilenanzahl

            ' sortieren nach KWS-BMK
            ' nach Einbauort SPS Rack (Stationsnummer)
            ' Daten sortieren
            SortTable tabelleDaten, SpalteAnlage, Sortierspalte2, Sortierspalte
    

            ' ab hier Seitenzahlen vergeben
            ' sowie sortiert vergebeben, dabei berücksichtigen ob Pneumatik oder nicht
            KennzeichenOld = "Leerplatz"

            For i = 3 To zeilenanzahl
                If .Cells.Item(i, SpalteAnlage) <> KennzeichenOld Then
                    SeitePneumatik = 1
                    Seite = 1
                End If
                KennzeichenOld = .Cells.Item(i, SpalteAnlage)
                ' Prüfen ob nicht Segmentvorlage ohne Seite
                If .Cells.Item(i, SpalteSegmentvorlage) <> "Sensor_ohne_SLP" Then
                    If .Cells.Item(i, Sortierspalte2) <> vbNullString Then ' prüfen ob Sortierspalte nicht leer
                        If .Cells.Item(i, SpaltePneumatik) = vbNullString Then
                            .Cells.Item(i, SpalteSeitenzahl) = Seite
                            Seite = Seite + 1
                        ElseIf .Cells.Item(i, SpaltePneumatik) <> vbNullString Then
                            .Cells.Item(i, SpalteSeitenzahl) = SeitePneumatik
                            SeitePneumatik = SeitePneumatik + 1
                        Else
                            .Cells.Item(i, SpalteSeitenzahl) = vbNullString
                        End If
                    Else
                        .Cells.Item(i, SpalteSeitenzahl) = vbNullString
                    End If
                Else
                    .Cells.Item(i, SpalteSeitenzahl) = vbNullString
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

