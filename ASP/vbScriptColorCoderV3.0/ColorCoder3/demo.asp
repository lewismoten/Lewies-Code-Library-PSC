<!--#INCLUDE FILE="clsColorVBS.asp"-->
<TITLE>Lewies Color Coder Version 3</TITLE>
<BODY bgColor="47659A" text="white" link="white" vlink="white" style="font-family:Arial;">
<H1>Lewies Color Coder</H1>
<H4>version 3.0</H4>
<FORM method="post">
Enter the vbScript source code:<BR>
<TEXTAREA name="VBS" cols="50" rows="10"><%=Server.HTMLEncode(Request.Form("VBS"))%></TEXTAREA><BR>
<INPUT type="Submit" value="Colorize">
</FORM>
<CENTER>
	Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
</CENTER>
<%
If Not Request.Form("VBS") = "" Then

	Response.Write "<HR>"
	
	Set ColorVBS = New clsColorVBS
	ColorVBS.VBS = Request.Form("VBS")
	Response.Write ColorVBS.HTML()
	Set ColorVBS = Nothing
	
End If
%>
