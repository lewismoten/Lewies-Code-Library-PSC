<%
Dim objConn
Dim objRs
Dim strSQL
Dim strConnection
Dim strHTML

' If you know the exact physical path to your access database,
' then modify this line to point to the database.
strConnection = _
	"DRIVER=Microsoft Access Driver (*.mdb);" & _
	"DBQ=" & Server.MapPath("news.mdb")

strSQL = "SELECT [Title], [Content] FROM [News] WHERE [NewsID] = " & Request.QueryString("NewsID")

Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")

objConn.Open strConnection

Set objRs = objConn.Execute(strSQL)

Response.Write "<H1>"
Response.Write Server.HTMLEncode(objRs.Fields("Title").Value & "")
Response.Write "</H1>"

Response.Write "<P>"
Response.Write objRs.Fields("Content").Value & ""
Response.Write "</P>"

Response.Write "[<A href=""FrontPage.asp"">Font Page</A>]"
Set objRs = objConn.Execute(strSQL)

Set objRs = Nothing

objConn.Close

Set objConn = Nothing
%>