<%
Dim oConn
Dim oRs
Dim sSQL
Dim nFileID

nFileID = Request.QueryString("FileID")

If Not nFileID = "" And IsNumeric(nFileID) Then

	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRs = Server.CreateObject("ADODB.Recordset")

	oConn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\Upload.mdb"

	sSQL = "SELECT FileName, ContentType, BinaryData FROM Files WHERE FileID = " & Request.QueryString("FileID")

	oRs.Open sSQL, oConn, 3, 3

	If Not oRs.EOF Then
		Response.ContentType = oRs(1)
		Response.BinaryWrite oRs(2)
	Else
		Response.Write("File could not be found")
	End If

	oRs.Close
	oConn.Close

	Set oRs = Nothing
	Set oConn = Nothing
Else
	Response.Write("File could not be found")
End If
%>

