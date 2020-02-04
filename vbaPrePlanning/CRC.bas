Attribute VB_Name = "CRC"
Option Explicit
 
' Polynom-Tabelle
Dim bCRC32Init As Boolean
Dim nCRC32LookUp() As Long

Public Sub CRC32_Init()
    ' Polynom-Tabelle erstellen
    ' Hier wird das offizielle Polynom verwendet, das
    ' auch von WinZip/PKZip verwendet wird
 
    ' Falls die LookUp-Tabelle bereits erstellt...
    If bCRC32Init Then Exit Sub
 
    Const nPolynom = &HEDB88320
 
    Dim i As Long
    Dim u As Long
 
    ReDim nCRC32LookUp(255)
    Dim nCRC32 As Long
 
    For i = 0 To 255
        nCRC32 = i
        For u = 0 To 7
            If (nCRC32 And 1) Then
                nCRC32 = (((nCRC32 And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) _
        Xor nPolynom
            Else
                nCRC32 = ((nCRC32 And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next u
        nCRC32LookUp(i) = nCRC32
    Next i
    bCRC32Init = True
End Sub

' Der optionale Parameter "nResult" sollte nur von
' CRC32_File verwendet werden!
Public Function CRC32(ByRef Bytes() As Byte, _
                      Optional ByVal nResult As Long = &HFFFFFFFF) As Long
 
    Dim i As Long
    Dim Index As Long
    Dim nSize As Long
 
    ' ggf. LookUp-Tabelle erstellen...
    If Not bCRC32Init Then CRC32_Init
 
    nSize = UBound(Bytes)
    For i = 0 To nSize
        Index = (nResult And &HFF) Xor Bytes(i)
        nResult = (((nResult And &HFFFFFF00) \ &H100) And 16777215) _
        Xor nCRC32LookUp(Index)
    Next i
 
    CRC32 = Not (nResult)
End Function

' CRC32-Checksumme einer Datei berechnen
Public Function CRC32FromFile(ByVal sFile As String) As Long
    ' Um die Verarbeitung von großen Dateien zu beschleunigen,
    ' wird der Inhalt blockweise ausgelesen. Hierbei hat sich
    ' eine Blockgröße von 4096 Bytes (4 KB) als sehr gut erwiesen
    Const BlockSize As Long = 4096
 
    Dim FileSize As Long
    Dim FilePos As Long
    Dim BytesToRead As Long
    Dim nResult As Long
    Dim Bytes() As Byte
    Dim F As Integer
 
    On Error GoTo ErrHandler
 
    ' Datei binär öffnen
    F = FreeFile
    Open sFile For Binary Access Read Shared As #F
 
    ' Dateigröße
    FileSize = LOF(F)
 
    ' Datei blockweise einlesen und verarbeiten
    nResult = &HFFFFFFFF
    ReDim Bytes(BlockSize - 1)
    While FilePos < FileSize
        If FilePos + BlockSize > FileSize Then
            BytesToRead = FileSize - FilePos
            ReDim Bytes(BytesToRead - 1)
        Else
            BytesToRead = BlockSize
        End If
 
        Get #F, , Bytes()
        nResult = Not (CRC32(Bytes, nResult))
 
        FilePos = FilePos + BytesToRead
    Wend
    Close #F
 
    CRC32FromFile = Not (nResult)
    On Error GoTo 0
    Exit Function
 
ErrHandler:
    If F > 0 Then Close #F
    CRC32FromFile = -1
End Function

' Skript zur Ermittlung der CRC-Summe einer Zeile
' V3.1
'20.01.2020
'angepasst für MH04
'
' Christian Langrock + TH
' christian.langrock@actemium.de


Public Sub CRC_Zeile()

    Dim wkb As Workbook
    Dim ws1 As Worksheet
    Dim tabelleDaten As String
    Dim zeilenanzahl As Integer
    Dim i As Integer
    Dim y As Integer
    Dim nCRCSum As Long
    
    
    Dim sText As String
    Dim sTextGesamt As String
    
    
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
 
        ' Spaltenbreiten anpassen
        ActiveSheet.Columns.Item("BF").Select
        Selection.ColumnWidth = 15
        ActiveSheet.Columns.Item("BG").Select
        Selection.ColumnWidth = 15

        ActiveSheet.Columns.Item("BH").Select
        Selection.ColumnWidth = 12
        ActiveSheet.Columns.Item("BI").Select
        Selection.ColumnWidth = 12
 
        '*********** Checksumme von aktuell nach alt kopieren******************
 
        For i = 3 To zeilenanzahl
            Cells(i, "BG") = Cells(i, "BF")
        Next i
 
        '***********************************************************************

        ' CRC für jede Zelle berechnen
        For i = 3 To zeilenanzahl
            sTextGesamt = vbNullString                     ' neue Zeile Text löschen
            For y = 2 To 54
                sText = Cells(2, y) & .Cells(i, y)
                'MsgBox sText
                sTextGesamt = sTextGesamt & sText

                'MsgBox "CRC32-Checksumme: " & CStr(nCRCSum) & " bzw. &H" & Hex$(nCRCSum)
       
            Next y
            ' CRC schreiben
            nCRCSum = CRC32(StrConv(sTextGesamt, vbFromUnicode))
            .Cells(i, "BF") = "H" & Hex$(nCRCSum)
        Next i
 
 
        '*********** Checksumme vergleichen, markieren von Unterschieden und Datum erzeugen******************
        'Datum wird nur bei unterschiedlicher Checksumme neu generiert
        Dim Datum
        Datum = Format(Date, "dd.mm.yyyy")
        For i = 3 To zeilenanzahl
            If Cells(i, "BG") = Cells(i, "BF") Then
                Cells(i, "BG").Interior.ColorIndex = 4
                Cells(i, "BF").Interior.ColorIndex = 4
                Cells(i, "BI") = Cells(i, "BH")
            Else
                Cells(i, "BG").Interior.ColorIndex = 3
                Cells(i, "BF").Interior.ColorIndex = 3
                Cells(i, "BI") = Cells(i, "BH")
                Cells(i, "BH") = Datum
            End If
        Next i



        '***********************************************************************
     
    End With
End Sub




