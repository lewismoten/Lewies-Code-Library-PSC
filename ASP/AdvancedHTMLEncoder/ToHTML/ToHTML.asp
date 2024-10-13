<%
' Title: ToHTML
' Description: Formats a piece of text for HTML viewing.
' Author: Lewis Moten
' Email: lewis@moten.com
' URL: http://www.lewismoten.com
' Created: Saturday, May 05, 2001
' (c)Copyright 2001 Lewis Edward Moten III, All rights reserved.
' ------------------------------------------------------------------------------
Function ToHTML(ByRef pString)
	Dim lObjRegExp
	Dim lLngStart
	Dim lLngEnd
	
	If VarType(pString) = vbNull Then Exit Function
	
	ToHTML = pString
	
	' Parse TAGS
	Set lObjRegExp = New RegExp
	lObjRegExp.Global = True
	lObjRegExp.Pattern = "<[^>]*>"
	ToHTML = lObjRegExp.Replace(pString, "")
	Set lObjRegExp = Nothing

	' HTML Encoding
	ToHTML = Server.HTMLEncode(ToHTML)
	
	' Change Carriage Returns and Line Feeds to HTML
	ToHTML = Replace(ToHTML, vbCrLf, "<BR>")
	ToHTML = Replace(ToHTML, vbCr, "<BR>")
	ToHTML = Replace(ToHTML, vbLf, "<BR>")
	
	' Change Tabs to HTML
	ToHTML = Replace(ToHTML, vbTab, "&nbsp;&nbsp;&nbsp; ")
	
	' Change double-space to HTML
	While Not InStr(1, ToHTML, "  ") = 0
		ToHTML = Replace(ToHTML, "  ", "&nbsp; ")
	Wend
End Function
' ------------------------------------------------------------------------------
%>