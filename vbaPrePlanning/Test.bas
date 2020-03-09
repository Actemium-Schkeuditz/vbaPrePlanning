Attribute VB_Name = "Test"
'todo Erstellen von Auslesen-Funktion:

' Schritt 1: Daten von SPS Kanal 1 in SPS Kanal 2 schreiben:
'   Erledigt        - Steckplatz übertragen
'   Erledigt        - Kanal übertragen und um 1 erhöhen, da BMK zwei Funktionen hat (erledigt)
'   In Bearbeitung  - Adresse von "SPS-Adresse: Adresse [1]" übertragen in "SPS-Adresse: Adresse [2]" und um ein Bit erhöhen, da BMK zwei Funktionen hat

' Schritt 2: Daten von SPS Kanal 1 korrigieren:
'   Offen           - Korrektur der Kanalzählung vom nächsten BMK
'   Erledigt        - Korrektur der Adressen

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

    Dim spalteSig_1_Steckplatz As String
    Dim spalteSig_2_Steckplatz As String
    Dim spalteSig_1_Kanal As String
    Dim spalteSig_2_Kanal As String

    Dim SpalteTest As String


    ' Spalten definieren
    spalteKartentyp = "CA"                       ' Spalte: ACT.PLS.SIGNAL_1.KARTENTYP de_DE
    spalteSig_1_Steckplatz = "CC"                ' Spalte: ACT.PLS.SIGNAL_1.STECKPLATZ de_DE
    spalteSig_2_Steckplatz = "CQ"                ' Spalte: ACT.PLS.SIGNAL_2.STECKPLATZ de_DE
    spalteSig_1_Kanal = "CD"                     ' Spalte: ACT.PLS.SIGNAL_1.KANAL de_DE
    spalteSig_2_Kanal = "CR"                     ' Spalte: ACT.PLS.SIGNAL_2.KANAL de_DE
    SpalteAnlage = "B"

    Set myblatt = ThisWorkbook.Worksheets("EplSheet")
    Datei = ThisWorkbook.Name


    'Zaehler3 = myblatt.Cells(Rows.Count, "A").End(xlUp).Row + 1

    ''''---> Funktion bezogen auf die Zeilen 3 bis 300''''

    With Workbooks(Datei).Sheets("EplSheet")


        'Bit = Right(myblatt.Cells(13, 83).Value, 1)
        '=VERKETTEN(LINKS(CE13;LÄNGE(CE13)-1);RECHTS(CE13;1)+1)


        For Zaehler1 = 3 To .Cells(Rows.Count, SpalteAnlage).End(xlUp).Row 'Gültigkeit in Spalte "A" von Zeile 3 bis Ende der Spalte

            '''For Zaehler1 = 3 To 300  ' Gültigkeit von Zeile 3 bis 300
            '''For Zaehler2 = 3 To .Cells(Rows.Count, "CA").End(xlUp).Row  'Gültigkeit in Spalte "CA" von Zeile 3 bis Ende der Spalte
 
 
            '' Wenn die Beschreibungen von Spalte ACT.PLS.SIGNAL_1.KARTENTYP de_DE gleich "CPX 5/2 bistabil" und die Spalte "Anlage" nicht leer ist dann
            If myblatt.Cells(Zaehler1, spalteKartentyp).Value = "CPX 5/2 bistabil" And myblatt.Cells(Zaehler1, SpalteAnlage).Value <> 0 Then
    
       
                myblatt.Cells(Zaehler1, spalteSig_2_Steckplatz).Value = myblatt.Cells(Zaehler1, spalteSig_1_Steckplatz).Value 'Übernimm den Wert von Spalte ACT.PLS.SIGNAL_1.STECKPLATZ de_DE in Spalte ACT.PLS.SIGNAL_2.STECKPLATZ de_DE
                myblatt.Cells(Zaehler1, spalteSig_2_Kanal).Value = myblatt.Cells(Zaehler1, spalteSig_1_Kanal).Value + 1 'Übernimm den Wert von Spalte ACT.PLS.SIGNAL_1.KANAL de_DE und erhöhe den Wert um 1  in Spalte ACT.PLS.SIGNAL_2.KANAL de_DE
        
        
                Zaehler1 = Zaehler1 + 1



            End If
        Next
    End With

End Sub
