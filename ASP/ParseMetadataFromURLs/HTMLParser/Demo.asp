<!--#INCLUDE FILE="clsHTMLParser.asp"-->
<%
Dim StrURL
Dim StrHTML
Dim ObjParser

StrURL = Request.QueryString("URL")
%>
<H1>HTML Parser</H1>
<P>
	This script will request the page from the
	server specified in the URL and parse the
	Title, Description, and Keywords for you.
</P>
<FORM>
	<INPUT size="50" name="URL" value="<%=StrURL%>"><BR>
	<INPUT type="Submit" value="Parse">
</FORM>
Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.<BR><BR>
<%
If Not StrURL = "" Then
	Set ObjParser = New clsHTMLParser
	With ObjParser
		StrHTML = .GetURL(StrURL)
		%>
		<TABLE border="1">
			<TR>
				<TD>Title</TD>
				<TD><%=.Title%></TD>
			</TR>
			<TR>
				<TD>Keywords</TD>
				<TD><%=.Keywords%></TD>
			</TR>
			<TR>
				<TD>Description</TD>
				<TD><%=.Description%></TD>
			</TR>
		</TABLE>
		<HR>
		<%
		Response.Write Replace(Server.HTMLEncode(StrHTML), vbCrLf, "<BR>")
	End With
	Set ObjParser = Nothing
End If
%>