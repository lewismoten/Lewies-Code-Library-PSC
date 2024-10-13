<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
Dim objConn
Dim objRs
Dim objDB
Dim strSQL
Dim lngImageID

lngImageID = Request.QueryString("ImageID")

Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")

Set objDB = New clsDatabase
objConn.Open objDB.ConnectionString
Set objDB = Nothing

strSQL = _
	"SELECT " & _
		"[ImageData] " & _
	"FROM " & _
		"[Images] " & _
	"WHERE " & _
		"[ImageID] = " & lngImageID
			
objRs.Open strSQL, objConn, 3, 3

If Not objRs.EOF Then
	Response.BinaryWrite objRs(0)
Else
	Response.Write "Image not found."
End If

objRs.Close

objConn.Close

Set objRs = Nothing
Set objConn = Nothing
%>
