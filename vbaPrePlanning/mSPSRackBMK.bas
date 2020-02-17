Attribute VB_Name = "mSPSRackBMK"
' Skript zur Ermittlung der Anlagen und Ortskennzeichen der IO-Racks
' V0.4
' nicht fertig
' 07.02.2020
' angepasst f�r MH04
'
' Christian Langrock
' christian.langrock@actemium.de
'@folder(Kennzeichen.SPS_RACK)

Option Explicit

Public Sub SPS_RackBMK()
    ' Erzeugen des gesamten Anlagen und Ortskennzeichen f�r SPS-Rack
    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Long
    Dim i As Long
    Dim spalteStationsnummer As String
    
    Dim spalteEinbauortRack As String
    Dim spalteRackAnlagenkennzeichen As String
    'Dim spalteSPSKanal As String
    Dim spalteAnlagenkennzeichen As String
    Dim dataAnlagenkennzeichen As String
    Dim dataRackAnlagenkennzeichen As String
    Dim answer As Long
      
    ' Tabellen definieren
    tabelleDaten = "EplSheet"

    Set wkb = ActiveWorkbook
    Set ws1 = Worksheets.[_Default](tabelleDaten)
   
    Application.ScreenUpdating = False

    ' Tabelle mit Daten bearbeiten
    With ws1
   
        ' Herausfinden der Anzahl der Zeilen
        zeilenanzahl = .Cells.Item(Rows.Count, 2).End(xlUp).Row ' zweite Spalte wird gez�hlt
        'MsgBox zeilenanzahl

        spalteAnlagenkennzeichen = "B"
        spalteStationsnummer = "BU"
        spalteEinbauortRack = "BV"
        spalteRackAnlagenkennzeichen = "BW"
          
 
        answer = MsgBox("Spalte BU Stationsnummern und Einbauorte schon gepr�ft?", vbQuestion + vbYesNo + vbDefaultButton2, "Pr�fung der Daten")
        'Pr�fe Stationsnummer
        If answer = vbYes Then
 
            ' Spaltenbreiten anpassen
            ActiveSheet.Columns.Item(spalteRackAnlagenkennzeichen).Select
            Selection.ColumnWidth = 35
    
        
            ' Daten schreiben
            For i = 3 To zeilenanzahl
                ' lesen von Feld Anlagenkennzeichen, f�hrende Leerzeichen entfernen
                dataAnlagenkennzeichen = LTrim$(.Cells.Item(i, spalteAnlagenkennzeichen))
                ' Pr�fe ob Stationsnummer mit Eintrag
                If .Cells.Item(i, spalteStationsnummer) <> vbNullString Then
                    ' Anlagenkennzeichen ermitteln
                    dataRackAnlagenkennzeichen = "=" + Left$(dataAnlagenkennzeichen, InStr(1, dataAnlagenkennzeichen, "."))
                    If Len(.Cells.Item(i, spalteStationsnummer)) = 1 Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "A.S0" + .Cells.Item(i, spalteStationsnummer)
                    Else
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "A.S" + .Cells.Item(i, spalteStationsnummer)
                    End If
                    ' wenn Einbauort nicht leer
                    If .Cells.Item(i, spalteEinbauortRack) <> vbNullString Then
                        dataRackAnlagenkennzeichen = dataRackAnlagenkennzeichen + "+" + .Cells.Item(i, spalteEinbauortRack)
                    End If
                    ' Daten schreiben
                    .Cells.Item(i, spalteRackAnlagenkennzeichen) = dataRackAnlagenkennzeichen
                Else
                    ' Daten leeren
                    .Cells.Item(i, spalteRackAnlagenkennzeichen) = vbNullString
                End If
            Next i
        Else
            MsgBox "Bitte Skript Stationsnummer ausf�hren und pr�fen!"
        End If
    
    End With

End Sub