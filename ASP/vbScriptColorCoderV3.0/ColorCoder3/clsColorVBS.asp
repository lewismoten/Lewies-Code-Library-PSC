<%
Class clsColorVBS

	Public VBS
	Private KeyWords
	Public KeywordColor
	Public StringColor
	Public CommentColor

	Private Sub Class_Initialize()
		
		' These colors must be different from each other
		KeywordColor = "blue"
		StringColor = "red"
		CommentColor = "green"
		
		KeyWords = Array( _
			"And", "Array", "As", "Boolean", "ByRef", "Byte", "ByVal", _
			"Call", "Case", "CBool", "CByte", "CCur", "CDate", "CDbl", _
			"CDec", "Class", "Close", "CInt", "CLng", "Compare", _
			"Const", "CSng", "CStr", "Currency", "CVar", "CVErr", _
			"Date", "Decimal", "Declare", "DefBool", "DefByte", _
			"DefCur", "DefDate", "DefDbl", "DefInt", "DefLng", _
			"DefObj", "DefSng", "DefStr", "DefVar", "Dim", "Do", _
			"DoEvents", "Double", "Each", "Else", "ElseIf", "EndIf", _
			"Empty", "End", "Enum", "Eqv", "Erase", "Error", "Event", _
			"Exit", "Explicit", "False", "Fix", "For", "Friend", _
			"Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", _
			"Implements", "In", "Input", "InputB", "Int", "Integer", _
			"Is", "LBound", "Len", "LenB", "Let", "Like", "Local", _
			"Lock", "Long", "Loop", "LSet", "Me", "Mod", "New", _
			"Next", "Not", "Nothing", "Null", "Open", "Option", _
			"Optional", "Or", "On", "Public", "Print", "Private", _
			"Property", "PSet", "Put", "RaiseEvent", "ReDim", _
			"Resume", "Return", "Seek", "Select", "Set", "Sgn", _
			"Single", "Spc", "Static", "Stop", "String", "Sub", _
			"Tab", "Text", "Then", "To", "True", "Type", "UBound", _
			"UnLock", "Variant", "Wend", "While", "With", "Write" _
		)
	End Sub

	Public Function Load(FileName)
		Dim FSO
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(Filename) Then
			VBS = FSO.OpenTextFile(FileName, ForReading, False).ReadAll()
		Else
			VBS = FileName & " does not exist."
		End If
		Set FSO = Nothing
	End Function

	Public Function HTML()
		Dim lstrHTML
		Dim lobjRegExp
		Dim Keyword
		Dim OldHTML
		
		lstrHTML = VBS
		lstrHTML = Server.HTMLEncode(lstrHTML)
		
		Set lobjRegExp = New RegExp
		lobjRegExp.Global = True
		lobjRegExp.IgnoreCase = True
		lobjRegExp.Multiline = False
		
		' Strings
		lstrHTML = Replace(lstrHTML, "&quot;", Chr(222))
		lobjRegExp.Pattern = "(" & Chr(222) & "[^" & Chr(222) & "]*" & Chr(222) & ")"
		lstrHTML = lobjRegExp.Replace(lstrHTML, "<FONT color=" & StringColor & ">$1</FONT>")
		lstrHTML = Replace(lstrHTML, Chr(222), "&quot;")
		
		' Comments
		lobjRegExp.Pattern = "(<FONT color=" & StringColor & ">[^<']*)'([^<']*</FONT>)"
		lstrHTML = lobjRegExp.Replace(lstrHTML, "$1" & Chr(222) & "$2")
		lobjRegExp.Pattern = "('[^\n]*)(\r\n)"
		lstrHTML = lobjRegExp.Replace(lstrHTML, "<FONT color=" & CommentColor & ">$1</FONT>$2")
		lstrHTML = Replace(lstrHTML, Chr(222), "'")

		' Remove string colors in comments
		lobjRegExp.Pattern = "(<FONT color=" & CommentColor & ">[^<]*)<FONT color=" & StringColor & ">([^<]*)</FONT>([^<]*</FONT>)"
		lstrHTML = lobjRegExp.Replace(lstrHTML, "$1$2$3")
		
		' Hilight VBs Reserved Keywords
		For Each Keyword In KeyWords
			lobjRegExp.Pattern = "\b" & Keyword & "\b"
			lstrHTML = lobjRegExp.Replace(lstrHTML, Chr(222) & Keyword & Chr(223))
		Next
		lstrHTML = Replace(lstrHTML, Chr(222), "<FONT color=" & KeywordColor & "><B>")
		lstrHTML = Replace(lstrHTML, Chr(223), "</B></FONT>")

		' Un-higlight in strings
		lstrHTML = Replace(lstrHTML, "<FONT color=" & StringColor & ">", Chr(222))
		lobjRegExp.Pattern = "(" & Chr(222) & "[^<]*)<FONT[^>]*><B>([^<]*)</B></FONT>"
		Do
			OldHTML = lstrHTML
			lstrHTML = lobjRegExp.Replace(lstrHTML, "$1$2")
		Loop While Not OldHTML = lstrHTML
		lstrHTML = Replace(lstrHTML, Chr(222), "<FONT color=" & StringColor & ">")
		
		' Un-higlight in comments
		lstrHTML = Replace(lstrHTML, "<FONT color=" & CommentColor & ">", Chr(222))
		lobjRegExp.Pattern = "(" & Chr(222) & "[^<]*)<FONT[^>]*><B>([^<]*)</B></FONT>"
		Do
			OldHTML = lstrHTML
			lstrHTML = lobjRegExp.Replace(lstrHTML, "$1$2")
		Loop While Not OldHTML = lstrHTML
		lstrHTML = Replace(lstrHTML, Chr(222), "<FONT color=" & CommentColor & ">")

		lstrHTML = Replace(lstrHTML, vbCrLf, "<BR>" & vbCrLf)
		lstrHTML = Replace(lstrHTML, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")

		HTML = _
		"<STYLE>" & _
		".vbScript{color:black;background-color:white;font-size:10pt;font-family:Courier New;}" & _
		".vbScript B{color:" & KeywordColor & "}" & _
		"</STYLE>" & _
		"<DIV class=vbScript>" & _
		"<FONT face=""Courier New"">" & _
		lstrHTML & _
		"</FONT>" & _
		"</DIV>"
	End Function
End Class
%>