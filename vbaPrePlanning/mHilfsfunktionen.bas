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



Public Function newExcelFile(ByVal sNewFileName As String, ByVal sFolder As String) As Boolean

Dim sFolderFile As String
Dim sConfigFolder As String
Dim wbnew As Workbook

sConfigFolder = "config"
sFolder = ThisWorkbook.Path & "\" & sConfigFolder & "\"
sFolderFile = ThisWorkbook.Path & "\" & sConfigFolder & "\" & sNewFileName

createFolder sFolder

If Dir(sFolderFile) = "" Then
    newExcelFile = True
Else
    newExcelFile = False
End If

If newExcelFile = True Then
    Set wbnew = Application.Workbooks.Add
    wbnew.SaveAs filename:=sFolderFile, FileFormat:=xlOpenXMLStrictWorkbook
    
    wbnew.Close
End If
End Function

Public Function fileExist(ByVal sFilename As String, ByVal sFolder As String) As Boolean

Dim sTestFile As String
sTestFile = ThisWorkbook.Path & "\" & sFolder & "\" & sFilename
    If Dir(sTestFile) <> "" Then
        'MsgBox "vorhanden"
        fileExist = True
    Else
        'MsgBox "nicht vorhanden"
        fileExist = False
    End If
End Function

Public Function createFolder(ByVal foldername As String) As Boolean
' create folder if not exist
If Dir(foldername, vbDirectory) = "" Then
  MkDir (foldername)
  createFolder = True
Else
  createFolder = False
End If
End Function




