<!--#INCLUDE FILE="ToHTML.asp"-->
<%
Dim StrHTML
StrHTML = Request.Form("HTML")
%>
<HTML>
	<HEAD>
		<TITLE>ToHTML Demonstration By Lewis Moten</TITLE>
	</HEAD>
	<BODY>
		<H1>ToHTML Demonstration</H1>
		<P>
			This code demonstrates how you can convert a string to
			HTML text format.  It goes beyond the power of the
			Server.HTMLEncode method in that it parses out all HTML
			tags, changes new-line characters to "BR" tags, and also
			formats double-spacing and tabs.
		</P>
		<FORM method="post">
			HTML To Format:<BR>
			<TEXTAREA rows="7" cols="40" name="HTML"><%=Server.HTMLEncode(StrHTML)%></TEXTAREA><BR>
			<INPUT type="submit" value="Format">
		</FORM>
		Created By <A href="http://www.lewismoten.com">Lewis Moten</A>.
		<%
		If Not StrHTML = "" Then
			Response.Write("<HR>")
			Response.Write(ToHTML(StrHTML))
		End If
		%>
	</BODY>
</HTML>