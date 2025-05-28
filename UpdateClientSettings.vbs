Option Explicit

Dim fso, wshShell, tempFolder, zipUrl, tempZip, robloxVersionsPath, latestVersionPath
Dim objXMLHTTP, objStream, shellApp, folder, items, item
Dim clientSettingsPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' URLs
zipUrl = "https://github.com/004fer/DesyncClientSettings/raw/refs/heads/main/DesyncClientSettings.zip"

' Temp zip location
tempFolder = wshShell.ExpandEnvironmentStrings("%TEMP%")
tempZip = tempFolder & "\DesyncClientSettings.zip"

' Roblox Versions path
robloxVersionsPath = wshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Roblox\Versions"

' Find the latest version folder starting with "version-"
latestVersionPath = ""
If fso.FolderExists(robloxVersionsPath) Then
    Dim folderList, folderItem
    Set folderList = fso.GetFolder(robloxVersionsPath).SubFolders

    Dim maxDate
    maxDate = #1/1/1970#

    For Each folderItem In folderList
        If LCase(Left(folderItem.Name, 8)) = "version-" Then
            If folderItem.DateLastModified > maxDate Then
                maxDate = folderItem.DateLastModified
                latestVersionPath = folderItem.Path
            End If
        End If
    Next
End If

If latestVersionPath = "" Then
    WScript.Echo "Roblox version folder not found."
    WScript.Quit 1
End If

clientSettingsPath = latestVersionPath & "\ClientSettings"

' Download ZIP
If Not DownloadFile(zipUrl, tempZip) Then
    WScript.Echo "Failed to download ZIP file."
    WScript.Quit 1
End If

' Delete existing ClientSettings folder if exists
If fso.FolderExists(clientSettingsPath) Then
    fso.DeleteFolder clientSettingsPath, True
End If

' Extract ZIP
If Not ExtractZip(tempZip, latestVersionPath) Then
    WScript.Echo "Failed to extract ZIP file."
    WScript.Quit 1
End If

' Hide ClientSettings folder and contents recursively
HideFolderRecursive clientSettingsPath

' Delete ZIP after extraction
On Error Resume Next
fso.DeleteFile tempZip
On Error GoTo 0

' Finished
'WScript.Echo "ClientSettings updated successfully."

WScript.Quit 0


'-----------------------
Function DownloadFile(url, savePath)
    On Error Resume Next
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.open "GET", url, False
    objXMLHTTP.send

    If objXMLHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 'binary
        objStream.Open
        objStream.Write objXMLHTTP.responseBody
        objStream.SaveToFile savePath, 2 'overwrite
        objStream.Close
        DownloadFile = True
    Else
        DownloadFile = False
    End If
    On Error GoTo 0
End Function

Function ExtractZip(zipFile, extractTo)
    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.NameSpace(zipFile)
    Set items = folder.Items

    Dim targetFolder
    Set targetFolder = shellApp.NameSpace(extractTo)
    If targetFolder Is Nothing Then
        ExtractZip = False
        Exit Function
    End If

    targetFolder.CopyHere items, 16 ' 16 = No UI

    ' Wait a bit for extraction to finish (roughly)
    Dim startTime
    startTime = Timer
    Do While Timer - startTime < 5 ' wait 5 seconds
        WScript.Sleep 200
    Loop

    ExtractZip = True
    On Error GoTo 0
End Function

Sub HideFolderRecursive(path)
    On Error Resume Next
    Dim folder, subFolder, file

    Set folder = fso.GetFolder(path)
    folder.Attributes = folder.Attributes Or 2 ' hidden

    For Each file In folder.Files
        file.Attributes = file.Attributes Or 2 ' hidden
    Next

    For Each subFolder In folder.SubFolders
        HideFolderRecursive subFolder.Path
    Next
    On Error GoTo 0
End Sub