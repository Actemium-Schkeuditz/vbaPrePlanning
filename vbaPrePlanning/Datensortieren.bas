Attribute VB_Name = "Datensortieren"
Option Explicit
' Skript zum sortieren der Datensätze
' V0.2
' 04.02.2020
' erste Funktion getestet, weitere Filter müssen noch rein
' Christian Langrock
' christian.langrock@actemium.de

' ToDO: weitere Filter müssen noch rein
Public Sub sortieren()

 Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
      
      On Error GoTo ErrorHandle
      
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

' sortieren nach KWS-BMK
' nach Eibauort SPS Rack

    Dim Sortierspalte As String
    Dim Bereich As String
    Bereich = "A3:EZ10000"
    Sortierspalte = "B" ' sortieren nach KWS-BMK
    ActiveSheet.Range(Bereich).Sort _
     Key1:=Range(Sortierspalte & "3"), Order1:=xlAscending, _
     Key2:=Range("BW" & "3"), Order1:=xlAscending, _
     Header:=xlNo, MatchCase:=False, _
     Orientation:=xlTopToBottom

End With


BeforeExit:
   ' Set rCell = Nothing
    'Set rTable = Nothing
    Exit Sub
ErrorHandle:
    MsgBox Err.Description & " Procedure MakeProductCollection.", vbCritical, "Error"
    Resume BeforeExit
End Sub


