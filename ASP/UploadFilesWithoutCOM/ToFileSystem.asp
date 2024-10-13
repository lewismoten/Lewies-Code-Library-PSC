<!--#INCLUDE FILE="clsUpload.asp"-->
<%
Dim oUpload
Dim oFile
Dim sFileName
Dim oFSO
Dim sPath
Dim sNewData
Dim nLength
Dim bytBinaryData
Dim sPath

Const nForReading = 1
Const nForWriting = 2
Const nForAppending = 8

' grab the uploaded file data
Set oUpload = New clsUpload
Set oFile = oUpload("File1")

' parse the file name
sFileName = oFile.FileName
If Not InStr(sFileName, "\") = 0 Then
	sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
End If

' Convert the binary data to Ascii
bytBinaryData = oFile.BinaryData
nLength = LenB(bytBinaryData)
For nIndex = 1 To nLength
	sNewData = sNewData & Chr(AscB(MidB(bytBinaryData, nIndex, 1)))
Next

' Save the file to the file system
sPath = Server.MapPath(".\Uploads") & "\"
Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
oFSO.OpenTextFile(sPath & sFileName, nForWriting, True).Write sNewData
Set oFSO = Nothing

Set oFile = Nothing
Set oUpload = Nothing
%>
File has been saved in file system.  View this file:<BR><BR>
<A href="Uploads\<%=sFileName%>">Uploads\<%=sFileName%></A>

