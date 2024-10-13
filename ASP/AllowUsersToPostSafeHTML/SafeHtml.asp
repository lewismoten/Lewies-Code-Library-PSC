<%
'**************************************
    ' for :Allow users to post "Safe" HTML
    '**************************************
    '(c)Copyright 2001 Lewis Edward Moten III, All rights reserved.
    '**************************************
    ' Name: Allow users to post "Safe" HTML
    ' Description:This code pulls out all th
    '     e nasty tags that a user sholdn't use wh
    '     en posting content. It also pulls out an
    '     y javascript events assigned to any tags
    '     . A must have if you allow people to pos
    '     t HTML on your site.
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
    
    Function SafeHTML(ByVal pStrHTML)
    	
    	Dim lObjRegExp
    	If VarType(pStrHTML) = vbNull Then Exit Function
    	If pStrHTML = "" Then Exit Function
    	Set lObjRegExp = New RegExp
    	lObjRegExp.Global = True
    	lObjRegExp.IgnoreCase = True
    	lObjRegExp.Pattern = "<(/)?SCRIPT|META|STYLE([^>]*)>"
    	pStrHTML = lObjRegExp.Replace(pStrHTML, "&lt;$1SCRIPT$3&gt;")
    	lObjRegExp.Pattern = "<(/)?(LINK|IFRAME|FRAMESET|FRAME|APPLET|OBJECT)([^>]*)>"
    	pStrHTML = lObjRegExp.Replace(pStrHTML, "&lt;$1LINK$3&gt;")
    	lObjRegExp.Pattern = "(<A[^>]+href\s?=\s?""?javascript:)[^""]*(""[^>]+>)"
    	pStrHTML = lObjRegExp.Replace(pStrHTML, "$1//protected$2")
    	lObjRegExp.Pattern = "(<IMG[^>]+src\s?=\s?""?javascript:)[^""]*(""[^>]+>)"
    	pStrHTML = lObjRegExp.Replace(pStrHTML, "$1//protected$2")
    	lObjRegExp.Pattern = "<([^>]*) on[^=\s]+\s?=\s?([^>]*)>"
    	pStrHTML = lObjRegExp.Replace(pStrHTML, "<$1$3>")
    	Set lObjRegExp = Nothing
    	
    	SafeHTML = pStrHTML
    	
    End Function

%>