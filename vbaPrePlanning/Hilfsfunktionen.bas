Attribute VB_Name = "Hilfsfunktionen"
Option Explicit
Public Function SpaltenBuchstaben2Int(pSpalte As String) As Integer
    'ermittel der Spaltennummer aus den Spaltenbuchstaben
    SpaltenBuchstaben2Int = Columns(pSpalte).Column


End Function

