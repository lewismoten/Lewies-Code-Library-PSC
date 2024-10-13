<%
    '**************************************
    ' for :Server Timezone Offset
    '**************************************
    '(C) Lewis Moten 2000, all rights reserved.
    '**************************************
    ' Name: Server Timezone Offset
    ' Description:Returns the hour differenc
    '     e of the servers time from GMT.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:Hour representation of time zo
    '     ne offset from GMT
    '
    'Assumes:User is a little familliar with
    '     javaScript - but not required unless adj
    '     usting the value returned.
    '
    'Side Effects:I found that JavaScript di
    '     dn't return the correct number of minute
    '     s on both server script and client side 
    '     script (Both IE and Netscape). (Wierd - 
    '     I must be missing something). So I had t
    '     o add some code to turn the value return
    '     ed into the correct number. I have not h
    '     eard of anyone else with this problem be
    '     fore because I am not aware of too many 
    '     people programming in javascript with AS
    '     P. If you don't have this problem, then 
    '     I would like to know what version of Jav
    '     aScript you are running on your server. 
    '     Mine is JavaScript 5.5 build 5207
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
%>    
<SCRIPT language="javascript" runat="server">
    	/*
    		vbScript does not currently support timezone offset on the
    		current system time. JavaScript is used to assist in time
    		calculations on users browser against there time.
    		
    		however, it seems that the timezone offset does not fully
    		work for ASP. You may have to adjust the difference if
    		it does not appear to work properly.
    		
    		this function returns the difference in hours to
    		(GMT) Greenwich mean time
    		
    		DO NOT FLAT lnOffset!! Some timezones are offset by minutes
    		e.g. Kathmandu is GMT +05:45 so it would be 5.75 hours
    	*/
    	function TimezoneOffset()
    	{
    		// get current time
    		var ldNow = new Date();
    		
    		// get timezone - and convert to hours
    		var lnOffset = ldNow.getTimezoneOffset() / 60
    		
    		// javascript bug correction?? - What's up with this?
    		lnOffset = -(lnOffset +1)
    		
    		// return the offset
    		return(lnOffset);
    	}
</SCRIPT>