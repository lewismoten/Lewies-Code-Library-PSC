<%
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT
' WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
' OF MERCHANTABILITY AND/OR FITNESS FOR A  PARTICULAR
' PURPOSE.

' Copyright 2000. All rights reserved.
' Lewis Moten
' http://www.lewismoten.com/
' email: lewis@moten.com

' if this string is found withing the whois replay, then it is assumed that the
' domain name is available.  The default value is "no match"
Const strAvailableString = "NO MATCH"

' This is the text that represents the end of a copyright message.  If
' "RemoveCopyright" is enabled, then all previous text will be removed.
' the default value is "policy."
Const strEndCopyright = "policy."

' Query is the actual domain that is sent for a whois lookup.  If this is
' not specified, then the whois lookup will submit a blank query.
Dim strQuery
strQuery = Request.QueryString("Query")

' Timeout is the length of time measured in seconds before the query
' will timeout and allow your script to continue to run.  Default
' value is 10 seconds
Const lngTimeout = 10

' The server address is the location of the whois database on the
' internet.  It may be an IP address or a domain name.  Default value
' is whois.networksolutions.com
Dim strWhoisServerAddress
strWhoisServerAddress = Request.QueryString("WhoisServerAddress")
If strWhoisServerAddress = "" Then strWhoisServerAddress = "whois.networksolutions.com"

' The server port is the port in wich to connect.  The default value
' is 43
Dim strWhoisServerPort
strWhoisServerPort = Request.QueryString("WhoisServerPort")
If strWhoisServerPort = "" Then strWhoisServerPort = 43

' The remove copyright attribute allows you to remove any copyright
' text that appears before the "EndCopyright" property.
Dim bolRemoveCopyright
bolRemoveCopyright = Request.QueryString("RemoveCopyright") = "True"

%>
<H1>Advanced Whois</H1>
<FORM id=form2 name=form2>
	<TABLE>
		<TR>
			<TD>Domain</TD>
			<TD><INPUT name="Query" value="<%=strQuery%>"></TD>
		</TR>
		<TR>
			<TD>Whois Server Address</TD>
			<TD><INPUT name="WhoisServerAddress" value="<%=strWhoisServerAddress%>"></TD>
		</TR>
		<TR>
			<TD>Whois Server Port</TD>
			<TD><INPUT name="WhoisServerPort" value="<%=strWhoisServerPort%>"></TD>
		</TR>
		<TR>
			<TD align="right">
				<INPUT type="checkbox" name="RemoveCopyright" value="True"<%If bolRemoveCopyright Then Response.Write(" checked")%>>
			</TD>
			<TD>Remove Copyright</TD>
		</TR>
	</TABLE>
	<INPUT type="submit" value="Lookup" id=submit1 name=submit1>
</FORM>
<HR>
<%
Dim objWhois

Set objWhois = Server.CreateObject("LewisMoten.Whois")

objWhois.AvailableString = strAvailableString
objWhois.EndCopyright = strEndCopyright
objWhois.RemoveCopyright = bolRemoveCopyright
objWhois.Timeout = lngTimeout
objWhois.WhoisServerAddress = strWhoisServerAddress
objWhois.WhoisServerPort = strWhoisServerPort

objWhois.Query strQuery

' Let user know if the domain is available.
Response.Write("The domain name <B>" & strQuery & "</B>")
If objWhois.IsAvailable Then
	Response.Write(" is ")
Else
	Response.Write(" is not ")
End If
Response.Write("available.<BR><BR>")

' Let user know where search was performed.
Response.Write("The following results were returned from ")
Response.Write(strWhoisServerAddress & ":" & strWhoisServerPort & ".<HR>")

' Show the Whois Results
Response.Write objWhois.WhoisHtml

Set objWhois = Nothing
%>
