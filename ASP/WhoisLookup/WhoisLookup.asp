<%@ Language=VBScript %>
<%
    '**************************************
    ' Name: Whois Lookup
    ' Description:Shows how you can do a Who
    '     Is lookup using a COM object capable of 
    '     using winsock.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
%>    
    <H1>Whois Lookup</H1>
    <FORM>
    	<INPUT type="text" name="DNS" value="<%=Request.QueryString("DNS")%>">
    	<INPUT type="submit" value="Lookup">
    </FORM>
    <HR>
    <%
    Dim objASPSocks
    Dim strResponse
    ' Goto ServerObjects.com for Eval Versio
    '     n.
    Set objASPSocks = Server.CreateObject("AspSock.Conn")
    objAspSocks.RemoteHost = "whois.networksolutions.com"
    objAspSocks.Port = 43
    objAspSocks.Open
    objAspSocks.WriteLn(Request.QueryString("DNS"))
    strResponse = objAspSocks.ReadBytesAsString(-1)
    objAspSocks.Close
    Set objASPSocks = Nothing
    	
    Response.Write "<PRE>" & strResponse & "</PRE>"
    %>