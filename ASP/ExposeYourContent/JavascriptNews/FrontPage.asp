<TABLE width="100%">
<TR>
<TD valign="top">
<%
Dim objConn
Dim objRs
Dim strSQL
Dim strConnection
Dim strHTML
Dim strLink


' If you know the exact physical path to your access database,
' then modify this line to point to the database.
strConnection = _
	"DRIVER=Microsoft Access Driver (*.mdb);" & _
	"DBQ=" & Server.MapPath("news.mdb")

strSQL = "SELECT TOP 10 [NewsID], [Title] FROM [News] ORDER BY [DatePosted] DESC"

Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")

objConn.Open strConnection

Set objRs = objConn.Execute(strSQL)

Response.Write "<H1>Front Page</H1>"

While Not objRs.EOF
	
	Response.Write "<P>" 

	' Open Link	
	Response.Write "<A href=""news.asp?NewsID=" & objRs.Fields("NewsID").Value & """>"

	' Write Title
	Response.Write Server.HTMLEncode(objRs.Fields("Title").Value & "")
	
	' Close Link
	Response.Write "</A>"

	Response.Write "</P>"
	
	objRs.MoveNext

Wend


Set objRs = Nothing

objConn.Close

Set objConn = Nothing
%>
</TD>
<TD valign="top" width="200">
<!--<SCRIPT src="http://www.otherdomain.com/news.js.asp"></SCRIPT>-->
<SCRIPT src="news.js.asp"></SCRIPT>
</TD>
</TABLE>
