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

Public Function fileExist(ByVal sfilename As String, ByVal sfolder As String) As Boolean

Dim sTestFile As String
sTestFile = ThisWorkbook.path & "\" & sfolder & "\" & sfilename
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
    'todo works not fine
    Dim sFolderFile As String
    Dim sfolder As String
    Dim sConfigFolder As String
    Dim wbnew As Workbook
    Dim closedBook As Workbook

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


Function ReadXmlPLCconfig(sfolder As String, sfilename As String) As cPLCconfig
    Dim sFile As String
    
    sFile = ThisWorkbook.path & "\" & sfolder & "\" & sfilename
  
    Dim xmlObj As Object
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
 
    xmlObj.async = False
    xmlObj.validateOnParse = False
    xmlObj.Load (sFile)
 
    Dim nodesThatMatter As Object
    Dim node            As Object
    
    Dim rdata As New cPLCconfig
    Dim sdata As New cPLCconfigData
    
    Set nodesThatMatter = xmlObj.SelectNodes("//PLCconfig")
   ' For Each node In nodesThatMatter
   '     'Task 1 -> print the XML file within the FootballInfo node:
   '     'Debug.Print node.XML
   '     Dim child   As Variant
   '     For Each child In node.ChildNodes
   '         'Task 2 -> print only the information of the clubs.  E.g. NorthClub, EastClub etc.
   '         'Debug.Print child.ChildNodes.Item(3).XML
   '     Next child
   ' Next node
    
    'Dim singleNode As Object
    'Set singleNode = xmlObj.SelectSingleNode("//PLCconfig/Station[@Number='1']")
    'Task 3 -> print only the node with number "1"
    'Debug.Print singleNode.XML
    Set sdata = Nothing
    Set rdata = Nothing
    
    Dim level1 As Object
    Dim level2 As Object
    Dim level3 As Object
    
    For Each level1 In nodesThatMatter
        For Each level2 In level1.ChildNodes
            Debug.Print level2.Attributes.getNamedItem("Number").NodeValue 'stationnumber
            sdata.Stationsnummer = level2.Attributes.getNamedItem("Number").NodeValue
            sdata.FirstInputAdress = level2.ChildNodes.Item(1).Text
            sdata.FirstOutputAdress = level2.ChildNodes.Item(2).Text
            
            
            For Each level3 In level2.ChildNodes.Item(3).ChildNodes
                Debug.Print level3.Attributes.Item(0).Text 'Kartentyp
                sdata.Kartentyp.Kartentyp = level3.Attributes.Item(0).Text 'Kartentyp
                Debug.Print level3.ChildNodes.Item(0).nodename
                Debug.Print level3.ChildNodes.Item(0).Text & vbCrLf 'ChannelsBeforSlot
                sdata.ReserveChannelsBefor = level3.ChildNodes.Item(0).Text
                Debug.Print level3.ChildNodes.Item(1).nodename
                Debug.Print level3.ChildNodes.Item(1).Text & vbCrLf 'ChannelsAfterSlot Value
                sdata.ReserveChannelsAfter = level3.ChildNodes.Item(1).Text
                Debug.Print level3.ChildNodes.Item(2).nodename
                Debug.Print level3.ChildNodes.Item(2).Text & vbCrLf 'ReserveChannelsPerSlot Value
                sdata.ReserveChannelPerSlot = level3.ChildNodes.Item(2).Text
                Debug.Print level3.ChildNodes.Item(3).nodename
                Debug.Print level3.ChildNodes.Item(3).Text & vbCrLf 'ReserveSlots Value
                sdata.ReserveSlot = level3.ChildNodes.Item(3).Text
            rdata.Addobj sdata
            Next
        Next
        
    Next
    
   Set ReadXmlPLCconfig = rdata
End Function

Public Function readXMLFile() As cPLCconfig
    ' read XML config File
    Dim sfolder As String
    Dim sFileNameConfig As String
    Dim bConfigFileIsNew As Boolean
    Dim bConfigFileExist As Boolean

    Dim rdata As New cPLCconfig
    
    sfolder = "config"
    sFileNameConfig = "PLC_Config.xml"
    
    
    ' Lege Datei an wenn nicht da
    bConfigFileIsNew = newExcelFile(sFileNameConfig, sfolder)

    ' wenn die Datei nicht angelegt wurde prüfe ob es diese schon gibt
    If bConfigFileIsNew = False Then
        'prüfe ob Datei vorhanden

        bConfigFileExist = fileExist(sFileNameConfig, sfolder)

        If bConfigFileExist Then
            'todo hier dann weiter wenn die Datei schon da ist
            'MsgBox "Datei gibt es schon"
            Set rdata = ReadXmlPLCconfig(sfolder, sFileNameConfig)
            'MsgBox Result
    
    
        ElseIf bConfigFileIsNew = True Then
            'todo hier weiter wenn Datei neu
        Else
            MsgBox "Fehler mit der SPSConfig.xlsx"
            rdata.Add 0, 0, "RESERVE"
        End If
    End If
    Set readXMLFile = rdata
End Function
