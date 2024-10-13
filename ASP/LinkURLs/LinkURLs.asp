<%
    '**************************************
    ' for :Link URLs
    '**************************************
    'Copyright (c) 2000, Lewis Moten. All rights reserved. Send all mofications to the author.
    '**************************************
    ' Name: Link URLs
    ' Description:Finds any URL found within
    '     specified text and creates a hyper link 
    '     for http, https, ftp, and email addresse
    '     s.
    ' By: Lewis Moten
    '
    '
    ' Inputs:asContent - Content to be parse
    '     d for URLs
    '
    ' Returns:returns the content with HTML 
    '     encoded hyperlinks.
    '
    'Assumes:This code assumes that you have
    '     vbScript 5.0 (or higher) installed on yo
    '     ur server. If not, you will instantly re
    '     ceive an error on your page.
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Function LinkURLs(ByRef asContent)
    	Dim loRegExp	' Regular Expression Object (Requires vbScript 5.0 and above)
    	
    	' If no content was received, exit the function
    	If asContent = "" Then Exit Function
    	
    	' Create Regular Expression object
    	Set loRegExp = New RegExp
    	
    	' Keep finding links after the first one.				
    	loRegExp.Global = True
    	
    	' Ignore upper/lower case
    	loRegExp.IgnoreCase = True
    	' Look for URLs
    	loRegExp.Pattern = "((ht|f)tps?://\S+[/]?[^\.])([\.]?.*)"
    	' Link URLs
    	LinkURLs = loRegExp.Replace(asContent, "<A href=""$1"">$1</A>$3")
    	' Look for email addresses
    	loRegExp.Pattern = "(\S+@\S+.\.\S\S\S?)"
    	' Link email addresses
    	LinkURLs = loRegExp.Replace(LinkURLs, "<A href=""mailto:$1"">$1</A>")
    	' Release regular expression object
    	Set oRegExp = Nothing
    	
    End Function
%>