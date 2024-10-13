<%
' IP Address Parser
'
'	Has ability to go out on the internet and determine what
'	your current IP address is.  Also allows you ability to
'	update namecheap.com with your current IP address.  Can
'	be modified to use with other dynamic dns services.
'
' Author: Lewis Moten
' Email: lewis@moten.com
' URL: http://www.lewismoten.com
'
' (C) Copyright 2003 Lewis Moten, All rights reserved.
'

' Different variations of XML objects
	'MSXML2.XMLHTTP.3.0
	'MSXML2.ServerXMLHTTP
	'Microsoft.XMLHTTP

Dim MyIP
Dim MyHost
Dim MyDomain
Dim MyPassword

Const AutoUpdate = True

MyIP = GetIP
MyDomain = Request.ServerVariables("SERVER_NAME")
MyHost = Mid(MyDomain, 1, InStr(1, MyDomain, ".") - 1)
MyDomain = Mid(MyDomain, InStr(1, MyDomain, ".") + 1)

MyPassword = "" ' enter your password here

If AutoUpdate Then

	' The following code automates updating your IP

	Session.Timeout = 20 ' minutes
	
	' Tell browser to refresh after N seconds (15 minutes = 15 * 60)
	Call Response.AddHeader("Refresh", 15 * 60)
	
	' If session IP address does not match
	If Not Session("HostIP") = MyIP Then
	
		' Notify client
		Response.Write "<B>"
		Response.Write "<FONT color=""red"">"
		Response.Write "IP Changed From " & Session("HostIP") & " to " & MyIP & "."
		Response.Write "</FONT>"
		Response.Write "</B>"
		
		' Update IP address
		Session("HostIP") = MyIP
		Call UpdateDNS(MyHost, MyDomain, MyPassword, MyIP)
		
	Else
	
		' IP Address has not changed
		
		Response.Write "IP " & MyIP & " has not changed."
	
	End If

	' Write timestamp to prevent confusion
	Response.Write "<HR>Timestamp: " & Now()

Else
' The following code manually updates your IP
%>
<P>
The Internet IP Address on this server is:<%=MyIP%>
</P>
<P>
	Fill out the following information to update
	your address on <A href="http://www.namecheap.com">namecheap.com</A>.
</P>
<FORM method="get" action="http://dynamicdns.<%=Domain%>/update">
	host: <INPUT name="host" value="<%=MyHost%>"><BR>
	domain: <INPUT name="domain" value="<%=MyDomain%>"><BR>
	password: <INPUT name="password" type="password"><BR>
	ip: <INPUT name="ip" value="<%=MyIP%>"><BR>
	<INPUT type="submit" value="update">
</FORM>
<%
End If
' ------------------------------------------------------------------------------
Public Function GetIP()

	Dim Spider
	Dim HTML
	Dim IP

	' Different variations of XML objects
	'Set Spider = Server.CreateObject ("MSXML2.XMLHTTP.3.0")
	'Set Spider = Server.CreateObject ("MSXML2.ServerXMLHTTP")
	Set Spider = Server.CreateObject ("Microsoft.XMLHTTP")
		
	' Could not create Internet Control
	If Err Then Exit Function

	With Spider
		.Open "GET", "http://checkip.dyndns.org/", False, "", ""
		.Send
		HTML = .ResponseText
	End With
	Set Spider = Nothing

	IP = Mid( HTML, 1, InStr(1, HTML, "<BR>", vbTextCompare) - 1)
	IP = Mid(IP, InStrRev(IP, ":") + 1)
	IP = Trim(IP)
	
	' Some people have mutliple for dynamic dns?
	If Not InStr(IP, ",") = 0 Then 
		IP = Mid(IP, 1, InStr(1, IP, ",") - 1)
		IP = Trim(IP)
	End If
		
	GetIP = IP
End Function
' ------------------------------------------------------------------------------
Public Sub UpdateDNS(Host, Domain, Password, IP)
	
	Dim Url
	
	Url = "http://dynamicdns." & Domain & "/update?"
	Url = Url & "host=" & Server.URLEncode(Host)
	Url = Url & "&domain=" & Server.URLEncode(Domain)
	Url = Url & "&password=" & Server.URLEncode(Password)
	Url = Url & "&ip=" & Server.URLEncode(IP)

	Set Spider = Server.CreateObject ("Microsoft.XMLHTTP")
	With Spider
		.Open "GET", Url, False, "", ""
		.Send
	End With
	Set Spider = Nothing

End Sub
' ------------------------------------------------------------------------------
%>
