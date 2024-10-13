<%
    '**************************************
    ' for :Strip HTML
    '**************************************
    'Copyright (C) 1999, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Strip HTML
    ' Description:Strips any HTML tags from 
    '     a string and returns the results. I use 
    '     this for data that users submit through 
    '     forms to prevent them from using HTML.
    ' By: Lewis Moten
    '
    '
    ' Inputs:asHTML - the string that may co
    '     ntain HTML code
    '
    ' Returns:Returns a string that has been
    '     stripped of HTML code.
    '
    'Assumes:This function assumes that you 
    '     have vbScript 5.0 or higher installed.
    '
    'Side Effects:Your page will come up wit
    '     h errors if you do not have vbScript 5.0
    '     or higher installed on your server (or o
    '     n the client browser if that is where it
    '     is being used)
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    ' Requires VBScript 5.0 installed on the
    '     server.
    Function StripHTML(ByRef asHTML)
    	Dim loRegExp	' Regular Expression Object
    	
    	' Create built in Regular Expression object
    	Set loRegExp = New RegExp
    	' Set the pattern to look for HTML tags
    	loRegExp.Pattern = "<[^>]*>"
    	
    	' Return the original string stripped of HTML
    	StripHTML = loRegExp.Replace(asHTML, "")
    	
    	' Release object from memory
    	Set loRegExp = Nothing
    End Function

%>