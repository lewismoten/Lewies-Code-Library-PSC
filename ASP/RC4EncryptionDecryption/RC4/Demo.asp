<%Option Explicit%>
<!--#INCLUDE FILE="RC4.asp"-->
<%
Dim lStrKey
Dim lStrMessage
Dim lStrResult
If Not Request.Form = "" Then
	lStrKey = Request.Form("Key")
	lStrMessage = Request.Form("Message")
	lStrResult = RC4(lStrMessage, lStrKey)
End If
%>
<HTML>
	<HEAD>
		<TITLE>RC4 Encryption</TITLE>
	</HEAD>
	<BODY>
		<H1>RC4 Encryption</H1>
		<P>
			This script encrypts and decrypts messages with the RC4
			algorithm.  Enter the key (password) and a message to
			encrypt or decrypt in the fields below.
		</P>
		<FORM method="Post">
			Key: <INPUT name="Key" value="<%=Server.HTMLEncode(lStrKey)%>"><BR>
			<BR>
			Message:<BR>
			<TEXTAREA name="Message" rows="6" cols="50"><%=Server.HTMLEncode(lStrResult)%></TEXTAREA>
			<BR>
			<INPUT type="Submit" value="Apply RC4">
		</FORM>
		<P>
			Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
			If you have questions about RCA and the algorithms
			that were released to the public domain, you can read the
			<A href="http://www.rsasecurity.com/solutions/developers/total-solution/faq.html">FAQ</A>.
		</P>
	</BODY>
</HTML>