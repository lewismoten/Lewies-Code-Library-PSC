<%@ Language=VBScript %>
<%Option Explicit%>
<!--#INCLUDE FILE="clsVBColor.asp"-->
<%
Dim StrCode
Dim ObjVBColor
Dim StrColorized

StrCode = Request.Form("Code")
If Not StrCode = "" Then
	Set ObjVBColor = New clsVBColor
	StrColorized = ObjVBColor.Colorize(StrCode)
	Set ObjVBColor = Nothing
End If
%>
<HTML>
	<HEAD>
		<TITLE>Lewies vbScript Colorizer Demo v2.0</TITLE>
		<LINK rel="stylesheet" type="text/css" href="vbColorCode.css">
	</HEAD>
	<BODY bgColor="#cccccc">
		<H1>Lewies vbScript Colorizer Demo v2.0</H1>
		<P>
			Copy and paste the vbScript that you wish to colorize in the
			form below and press submit.
		</P>
		<FORM method="post">
			<TEXTAREA name="code" cols="50" rows="10"><%=Server.HTMLEncode(StrCode)%></TEXTAREA><BR>
			<INPUT type="submit" value="Colorize!">
		</FORM>
		Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
		<HR>
		<NOBR><%=StrColorized%></NOBR>
	</BODY>
</HTML>
