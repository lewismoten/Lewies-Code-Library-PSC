<%
    '**************************************
    ' Name: Deny Website Copiers
    ' Description:Prevents web caching progr
    '     ams from accessing your website. Helps i
    '     ncrease performance of your website agai
    '     nst these programs that may cause DOS at
    '     tacks on websites with slow internet con
    '     nections - such as the upstream of a cab
    '     le modem.
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
    
    Call DenyWebsiteCopiers()
    Sub DenyWebsiteCopiers()
    	Dim lstrAgent
    	Dim lblnDenied
    	lblnDentied = False
    	lstrAgent = Request.ServerVariables("HTTP_USER_AGENT")
    	
    	' Product:		HTTrack Website Copier
    	' Website:		http://www.httrack.com/
    	' UserAgent:	Mozilla/4.5 (compatible; HTTrack 3.0x; Windows 98)
    	If InStr(1, lstrAgent, "HTTrack", vbTextCompare) Then lblnDenied = True
    	' Product:		Teleport Pro
    	' Website:		http://www.tenmax.com/teleport/pro/home.htm
    	' UserAgent:	Teleport Pro/1.29
    	If InStr(1, lstrAgent, "Teleport Pro", vbTextCompare) Then lblnDenied = True
    	
    	If lblnDenied Then
    		Response.Clear
    		%>
    		<P>
    			Some individuals have used programs to copy this entire website to there computer
    			for offline browsing. While this may be a large benefit to them, it causes a lot
    			of performance problems with this website and is considered a Denial of Service
    			attack.
    		</P>
    		<P>
    			The browser that you are using appears to be from such a program. Please use
    			a common browser such as Netscape or Microsoft. You will not only benefit from
    			acquireing content, but it will also be guarenteed to be the latest information
    			available from the website.
    		</P>
    		<%
    		Response.End
    	End If
    End Sub
    %>