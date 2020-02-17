Attribute VB_Name = "mHilfsfunktionen"
' Hilfsfunktionen die öfters benötigt werden
' V0.4
' 10.02.2020
' neu: SortTable
' Christian Langrock
' christian.langrock@actemium.de
'@folder Hilfsfunktionen
Option Explicit

Public Function SpaltenBuchstaben2Int(ByVal pSpalte As String) As Long
    'ermittel der Spaltennummer aus den Spaltenbuchstaben
    SpaltenBuchstaben2Int = Columns(pSpalte).Column


End Function

Public Sub SortTable(ByVal tablename As String, ByVal SortSpalte1 As String, ByVal SortSpalte2 As String, Optional ByVal SortSpalte3 As String)
    ' sortieren von Daten nach drei oder zwei Spalten
    ' Aufrufen der Tabele und Auswählen dieser
    ThisWorkbook.Worksheets(tablename).Activate
    ' Anzeige ausschalten
    Application.ScreenUpdating = False
        
    Dim rTable As Range
    Set rTable = Range("A3")                     ' Ohne die Überschriften
    If SortSpalte3 = vbNullString Then
        rTable.Sort _
        Key1:=Range(SortSpalte1 & "3"), Order1:=xlAscending, _
        Key2:=Range(SortSpalte2 & "3"), Order1:=xlAscending, _
        Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
    Else
        rTable.Sort _
        Key1:=Range(SortSpalte1 & "3"), Order1:=xlAscending, _
        Key2:=Range(SortSpalte2 & "3"), Order1:=xlAscending, _
        Key3:=Range(SortSpalte3 & "3"), Order1:=xlAscending, _
        Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom
    End If
        

    Set rTable = Nothing
    Exit Sub
End Sub



Public Function newExcelFile(ByVal sNewFileName As String, ByVal sfolder As String) As Boolean

Dim sFolderFile As String
Dim sConfigFolder As String
Dim wbnew As Workbook

sConfigFolder = "config"
sfolder = ThisWorkbook.path & "\" & sConfigFolder & "\"
sFolderFile = ThisWorkbook.path & "\" & sConfigFolder & "\" & sNewFileName

createFolder sfolder

newExcelFile = Dir(sFolderFile) = vbNullString

If newExcelFile = True Then
    Set wbnew = Application.Workbooks.Add
    wbnew.SaveAs filename:=sFolderFile, FileFormat:=xlOpenXMLStrictWorkbook
    
    wbnew.Close
End If
End Function

Public Function fileExist(ByVal sFilename As String, ByVal sfolder As String) As Boolean

Dim sTestFile As String
sTestFile = ThisWorkbook.path & "\" & sfolder & "\" & sFilename
    fileExist = Dir(sTestFile) <> vbNullString
End Function

Public Function createFolder(ByVal foldername As String) As Boolean
' create folder if not exist
If Dir(foldername, vbDirectory) = vbNullString Then
  MkDir (foldername)
  createFolder = True
Else
  createFolder = False
End If
End Function


Public Function ReadSecondExcelFile(ByVal sFilname As String, ByVal sfolder As String) As String
    '** Dimensionierung der Variablen
    Dim blatt As String
    Dim bereich As Range
    Dim zelle As Object
    Dim sFullFolder As String

    sFullFolder = ThisWorkbook.path & "\" & sfolder

    '** Angaben zur auszulesenden Zelle
    blatt = "Tabelle1"
    Set bereich = Range("A1")

    '** Bereich auslesen
    For Each zelle In bereich
        '** Zellen umwandeln
        zelle = zelle.Address(False, False)
        '** Eintragen in Bereich
        ReadSecondExcelFile = GetValue(sFullFolder, sFilname, blatt, zelle)
    Next zelle
 
End Function

Private Function GetValue(ByVal path As String, ByVal file As String, ByVal sheet As String, ByVal ref As String) As String
    Dim arg As String
    arg = "'" & path & "\[" & file & "]" & sheet & "'!" & Range(ref).Address(, , xlR1C1)
    GetValue = ExecuteExcel4Macro(arg)
End Function


Sub CopySheetFromClosedWB(sSheetname As String)
Application.ScreenUpdating = False
 
    Set closedBook = Workbooks.Open("D:\Dropbox\excel\articles\example.xlsm")
    closedBook.Sheets("Sheet1").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False
 
Application.ScreenUpdating = True
End Sub

Sub CopySheetToClosedWB(ByVal sNewFileName As String, sSheetname As String)
Application.ScreenUpdating = False
    
    Dim sFolderFile As String
    Dim sConfigFolder As String
    Dim wbnew As Workbook

    sConfigFolder = "config"
    sfolder = ThisWorkbook.path & "\" & sConfigFolder & "\"
    sFolderFile = ThisWorkbook.path & "\" & sConfigFolder & "\" & sNewFileName
    
    Set closedBook = Workbooks.Open(sFolderFile)
    Sheets(sSheetname).Copy Before:=closedBook.Sheets(sSheetname)
    closedBook.Close SaveChanges:=True
 
Application.ScreenUpdating = True
End Sub

Public Function ExtractNumber(ByVal str As String) As Long
    Dim i As Byte
    Dim ii As Byte
    ExtractNumber = 0
    'Prüfe ob Wert nicht leer

    If str <> vbNullString Then
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
        Exit For
        End If
    Next i
    For ii = i To Len(str)
        If Not IsNumeric(Mid(str, ii, 1)) Then
        Exit For
        End If
    Next ii
    ExtractNumber = Mid(str, i, Len(str) - (ii - i))
    Else
    ExtractNumber = 0
    End If
End Function

Public Function WorksheetExist(ByVal sWorksheetName As String, ByVal ws1 As Worksheet) As Boolean
 With ws1
   On Error Resume Next
   WorksheetExist = Worksheets(sWorksheetName).Index > 0
   End With
End Function
