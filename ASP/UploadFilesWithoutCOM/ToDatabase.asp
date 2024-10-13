<!--#INCLUDE FILE="clsUpload.asp"-->
<%
Dim oUpload
Dim oField
Dim oConn
Dim oRs
Dim sSQL
Dim sFileName

Set oUpload = New clsUpload
Set oFile = oUpload("File1")

' parse the file name
sFileName = oFile.FileName
If Not InStr(sFileName, "\") = 0 Then
	sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
End If

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.Recordset")

oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\Upload.mdb"

sSQL = "SELECT FileID, FileName, FileSize, ContentType, BinaryData FROM Files WHERE 1=2"

oRs.Open sSQL, oConn, 3, 3

oRs.AddNew

oRs.Fields("FileName") = sFileName
oRs.Fields("FileSize") = oFile.Length
oRs.Fields("ContentType") = oFile.ContentType
oRs.Fields("BinaryData").AppendChunk = oFile.BinaryData & ChrB(0)

oRs.Update

oRs.Close

sSQL = "SELECT Top 1 FileID, FileName From Files Order By FileID Desc"
oRs.Open sSQL, oConn

If Not oRs.EOF Then
	
	%>
	File has been saved in database.  View this file:<BR><BR>
	<A href="DataFile.asp?FileID=<%=oRs(0)%>"><%=oRs(1)%></A>
	<%
End If

Set oRs = Nothing
Set oConn = Nothing

Set oFile = Nothing
Set oUpload = Nothing
%>

