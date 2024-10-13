<%
    '**************************************
    ' Name: Site Post Check
    ' Description:Checks the HTTP Referer he
    '     ader to ensure people are not posting fr
    '     om other websites. You can include this 
    '     file if you use templates on your websit
    '     e, or just include it on the pages that 
    '     receive form data posts.
    'Warning - someone who knows there stuff can get around this by modifying there HOST file. this isn't 100% fool proof, but it may deter most from posting data from other websites.
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
    
    Call SitePostCheck()
    Sub SitePostCheck()
    	Dim lblnPost		' user posted data to page?
    	Dim lstrReferer		' page user is comming from
    	Dim lstrHost		' server user is on
    	lblnPost = Request.ServerVariables("REQUEST_METHOD") = "POST"
    	' if data wasn't posted, everythign is ok
    	If Not lblnPost Then Exit Sub
    	lstrReferer = Request.ServerVariables("HTTP_REFERER")
    	lstrHost = Request.ServerVariables("HTTP_HOST")
    	' If user is posting from antoher website
    	If InStr(1, lstrReferer, "//" & lstrHost & "/", vbTextCompare) = 0 Then
    		%>
    		<H1><FONT color="red">Security Alert</FONT></H1>
    		
    		<P>
    			The security of this web site does not allow you to post
    			data from other websites.
    		</P>
    		
    		<%
    		Response.End
    		
    	End If
    End Sub
    %>