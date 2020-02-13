Attribute VB_Name = "mConfigPLC"
' Skript zum Übertragen der SPS Konifguration nach Excel
' V0.1
' 13.02.2020
' neu
' Christian Langrock
' christian.langrock@actemium.de
'@folder (Daten.SPS-Konfig)
'todo Config Datei muss noch mit Werten befüllt werden

Option Explicit

Public Sub ConfigPLC()

Dim sFileNameConfig As String
Dim sFolder As String
Dim bConfigFileIsNew As Boolean
Dim bConfigFileExist As Boolean

sFolder = "config"
sFileNameConfig = "SPSConfig.xlsx"


' Lege Datei an wenn nicht da
bConfigFileIsNew = newExcelFile(sFileNameConfig, sFolder)

' wenn die Datei nicht angelegt wurde prüfe ob es diese schon gibt
If bConfigFileIsNew = False Then
'prüfe ob Datei vorhanden

bConfigFileExist = fileExist(sFileNameConfig, sFolder)

If bConfigFileExist Then
    'todo hier dann weiter wenn die Datei schon da ist
    MsgBox "Datei gibt es schon"
    ElseIf bConfigFileIsNew = True Then
    'todo hier weiter wenn Datei neu
    
End If

End If

End Sub
