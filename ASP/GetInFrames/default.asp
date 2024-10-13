<%
Dim sMainURL

sMainURL = "http://www.lewismoten.com/" ' default URL
If Not Request.QueryString("URL") = "" Then
	sMainURL = Request.QueryString("URL")
End If
%>
<FRAMESET rows="60,*">
	<FRAME src="title.htm">
	<FRAME src="<%=sMainURL%>">
</FRAMESET>
