Attribute VB_Name = "Test"
'todo Erstellen von Auslesen-Funktion:

' Schritt 1: Daten von SPS Kanal 1 in SPS Kanal 2 schreiben:
'   Erledigt        - Steckplatz übertragen
'   Erledigt        - Kanal übertragen und um 1 erhöhen, da BMK zwei Funktionen hat (erledigt)
'   In Bearbeitung  - Adresse von "SPS-Adresse: Adresse [1]" übertragen in "SPS-Adresse: Adresse [2]" und um ein Bit erhöhen, da BMK zwei Funktionen hat

' Schritt 2: Daten von SPS Kanal 1 korrigieren:
'   Offen           - Korrektur der Kanalzählung vom nächsten BMK

' Schritt 3: Daten von SPS Kanal 1 und SPS Kanal 2 in SPS Kanal 5 und SPS Kanal 6 schreiben:
'   Offen           - Adresse von Z01."SPS-Adresse: Adresse [1]" und Z01."SPS-Adresse: Adresse [2]" in Adresse von ES01."SPS-Adresse: Adresse [5]" und ES01."SPS-Adresse: Adresse [5]" schreiben

' Skript zur Ermittlung der SPS Kanäle



Public Sub CPX_5_2_bistabil()

Dim myblatt As Worksheet
Dim Zaehler1 As Integer
' Dim Zaehler2 As Integer
' Dim Zaehler3 As Integer
Dim Datei As Variant
Dim Bit As Variant






Set myblatt = ThisWorkbook.Worksheets("EplSheet")
Datei = ThisWorkbook.Name




'Zaehler3 = myblatt.Cells(Rows.Count, "A").End(xlUp).Row + 1

''''---> Funktion bezogen auf die Zeilen 3 bis 300''''

With Workbooks(Datei).Sheets("EplSheet")


'Bit = Right(myblatt.Cells(13, 83).Value, 1)
'=VERKETTEN(LINKS(CE13;LÄNGE(CE13)-1);RECHTS(CE13;1)+1)


  For Zaehler1 = 3 To .Cells(Rows.Count, "A").End(xlUp).Row  'Gültigkeit in Spalte "A" von Zeile 3 bis Ende der Spalte

' For Zaehler1 = 3 To 300  ' Gültigkeit von Zeile 3 bis 300
' For Zaehler2 = 3 To .Cells(Rows.Count, "CA").End(xlUp).Row  'Gültigkeit in Spalte "CA" von Zeile 3 bis Ende der Spalte
 
'' Wenn die Beschreibungen von Spalte CA bzw. Spalte 79 (ACT.PLS.SIGNAL_1.KARTENTYP de_DE) gleich "CPX 5/2 bistabil" und die Spalte A nicht leer ist dann

If myblatt.Cells(Zaehler1, 79).Value = "CPX 5/2 bistabil" And myblatt.Cells(Zaehler1, 1).Value <> 0 Then
    
    myblatt.Cells(Zaehler1, 95).Value = myblatt.Cells(Zaehler1, 81).Value        'Übernimm den Wert von Spalte 81 (ACT.PLS.SIGNAL_1.STECKPLATZ de_DE) in Spalte 95 (ACT.PLS.SIGNAL_2.STECKPLATZ de_DE)
    myblatt.Cells(Zaehler1, 96).Value = myblatt.Cells(Zaehler1, 82).Value + 1    'Übernimm den Wert von Spalte 82 (ACT.PLS.SIGNAL_1.KANAL de_DE) und erhöhe den Wert um 1  in Spalte 96 (ACT.PLS.SIGNAL_2.KANAL de_DE)`
    'myblatt.Cells(Zaehler1, 97).Value = myblatt.Cells(Zaehler1, 83).Value + 1    'Übernimm den Wert von Spalte 83 (SPS-Adresse: Adresse [1]) und erhöhe den Wert um 1 Bit in Spalte 96 (SPS-Adresse: Adresse [2])

        
    Zaehler1 = Zaehler1 + 1

'   myblatt.Cells(Zaehler3, 95).Value = myblatt.Cells(Zaehler2, 81).Value        'Übernimm den Wert von Spalte 81 (ACT.PLS.SIGNAL_1.STECKPLATZ de_DE) in Spalte 95 (ACT.PLS.SIGNAL_2.STECKPLATZ de_DE)
'   myblatt.Cells(Zaehler3, 96).Value = myblatt.Cells(Zaehler2, 82).Value + 1    'Übernimm den Wert von Spalte 82 (ACT.PLS.SIGNAL_1.KANAL de_DE) und erhöhe den Wert um 1  in Spalte 96 (ACT.PLS.SIGNAL_2.KANAL de_DE)
        
'   Zaehler3 = Zaehler3 + 1

End If
Next
End With

End Sub
