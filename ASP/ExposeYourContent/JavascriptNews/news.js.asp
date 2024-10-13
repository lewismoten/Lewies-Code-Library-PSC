<%
Dim objConn
Dim objRs
Dim strSQL
Dim strConnection
Dim strHTML
Dim strLink

' If you know the exact URL to your news page, replace these
' next three lines with the full URL.
strLink = Request.ServerVariables("PATH_INFO")
strLink = Replace(strLink, ".js", "")
strLink = "http://" & _
	Request.ServerVariables("SERVER_NAME") & _
	 strLink & _
	"?NewsID="

' If you know the exact physical path to your access database,
' then modify this line to point to the database.
strConnection = _
	"DRIVER=Microsoft Access Driver (*.mdb);" & _
	"DBQ=" & Server.MapPath("news.mdb")

strSQL = "SELECT TOP 5 [NewsID], [Title] FROM [News] ORDER BY [DatePosted] DESC"

Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")

objConn.Open strConnection

Set objRs = objConn.Execute(strSQL)

strHTML = "<B>Lastest News</B><BR>"

While Not objRs.EOF
	
	strHTML = strHTML & "<P>" 

	' Open Link	
	strHTML = strHTML & "<A href=""" & strLink
	strHTML = strHTML & objRs.Fields("NewsID").Value
	strHTML = strHTML & """>"

	' Write Title
	strHTML = strHTML & Server.HTMLEncode(objRs.Fields("Title").Value & "")
	
	' Close Link
	strHTML = strHTML & "</A>"

	strHTML = strHTML & "</P>"
	
	objRs.MoveNext
Wend

Response.Write "document.write('"

Response.Write ToJavascript(strHTML)

Response.Write "')"

Set objRs = Nothing

objConn.Close

Set objConn = Nothing

Function ToJavascript(ByVal pstrHTML)

	pstrHTML = Replace(pstrHTML, "\", "\\")
	pstrHTML = Replace(pstrHTML, vbCrLf, "\n")
	pstrHTML = Replace(pstrHTML, "'", "\'")

	ToJavascript = pstrHTML

End Function
%>