<!--
 THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT
 WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
 OF MERCHANTABILITY AND/OR FITNESS FOR A  PARTICULAR
 PURPOSE.

 Copyright 2000. All rights reserved.
 Lewis Moten
 http://www.lewismoten.com/
 email: lewis@moten.com
-->
<H1>Whois</H1>
<FORM>
	<INPUT name="Domain" value="<%=Request.QueryString("Domain")%>">
	<INPUT type="submit" value="Lookup">
</FORM>
<%
Dim objWhois
Set objWhois = Server.CreateObject("LewisMoten.Whois")
objWhois.Query Request.QueryString("Domain")
Response.Write(objWhois.WhoisHTML)
Set objWhois = Nothing
%>