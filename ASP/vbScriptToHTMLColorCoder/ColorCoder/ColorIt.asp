<%@ Language=VBScript %>
<%
Option Explicit
Dim sCode
Dim sColoredCode
sCode = Request.Form("Code")

If Not sCode = "" Then
	sColoredCode = ColorVBScript(sCode)
End If
%>
<HTML>
	<HEAD>
		<TITLE><%=Request.Form("Title")%></TITLE>
		<META http-equiv="Keywords" content="<%=Request.Form("Keywords")%>">
		<META name="description" content="<%=Request.Form("Description")%>">
		<SCRIPT language="javaScript" src="lookupWindow.js"></SCRIPT>
		<LINK rel="stylesheet" type="text/css" href="vbColorCode.css">
	</HEAD>
	<BODY>
		<TABLE>
			<TR>
				<TD>Title</TD>
				<TD><%=Request.Form("Title")%></TD>
			</TR>
			<TR>
				<TD>Description</TD>
				<TD><%=Request.Form("Description")%></TD>
			</TR>
		</TABLE>
		<HR>
		<%=sColoredCode%>
		<HR>
		<FONT size="1" face="arial">Created with Lewis Moten Color Code App</FONT>
	</BODY>
</HTML>
<%

' -----------------------------------------------------------------------------
Function ColorVBScript(ByRef NewCode)
	Dim lvReservedWords							' Array of Reserved Words in this langauge
	Dim lvConstants								' Array of recognized constants
	Dim lvFunctions								' Array of recognized functions
	Dim lvMethods								' Array of recognized methods
	Dim lvProperties							' Array of recognized properties
	Dim lvStatements							' Array of recognized statments
	Dim lvOperators								' Array of recognized operators
	Dim lvSeperators							' Array of single characters that seperate words.
	Dim lvLines									' Array of lines sent to this procedure
	
	Const lsSTRING_CHAR = """"					' Character that begins and ends strings.
	Const lsLINE_COMMENT_CHAR = "'"				' Character that begins a commented line
	Const lsLINE_COMMENT_CONTINUE_CHAR = "_"	' Character that continues comment to the next line
	
	Dim lsLine									' A single line within the code passed to the procedure
	Dim lsChar									' A single character within the current line being looked at.
	Dim lsString								' Contents of a string
	Dim lnStart									' Character Start Position that we are looking at.
	Dim lsWord									' A word being formed from the character being looked at.
	Dim lbBuildString							' Are we building a string?
	Dim lbBuildWord								' Are we building a word?
	Dim lsNewLine								' The new line that is being build from a line sent to this procedure.
	Dim lbBuildComment							' Are we building a comment?
	Dim lsResult								' The results to return when this procedure is finnished.
	
	
	' Define a list of characters that seperate words from next character.
	lvSeperators = Array(" ", vbTab, ":", "(", ")", ",", "'")

	' Define a list of Reserved Words
	lvReservedWords = Array( _
		"And", "Array", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", _
		"Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "Class", _
		"Close", "CInt", "CLng", "Const", "CSng", "CStr", "Currency", "CVar", _
		"Date", "Decimal", "Declare", "DefBool", "DefByte", "DefCur", _
		"DefDate", "DefDbl", "DefInt", "DefLng", "DefSng", "DefStr", _
		"DefVar", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "Empty", _
		"End", "Enum", "Eqv", "Erase", "Error", "Exit", "Explicit", "False", _
		"Fix", "For", "Friend", "Function", "Get", "Global", "If", "Imp", _
		"In", "Input", "InputB", "Int", "Integer", "Is", "LBound", "Len", _
		"LenB", "Local", "Lock", "Long", "Loop", "Me", "Mod", "New", "Next", _
		"Not", "Nothing", "Null", "Open", "Option", "Optional", "Or", "On", _
		"Public", "Print", "Private", "Property", "ReDim", "Resume", _
		"Select", "Set", "Sgn", "Single", "String", "Sub", "Then", "To", _
		"True", "Type", "UBound", "UnLock", "Variant", "Wend", "While", _
		"With", "Write", "Implements", "Compare", "Text", "Let" _
		)
		
	'lvReservedWords = Array()
	' Define a list of Constants
	lvConstants = Array( _
		"vbAbort", "vbAbortRetryIgnore", "vbApplicationModal", "vbArray", _
		"vbBinaryCompare", "vbBlack", "vbBlue", "vbBoolean", "vbByte", _
		"vbCancel", "vbCr", "vbCritical", "vbCrLf", "vbCurrency", _
		"vbCyan", "vbDatabaseCompare", "vbDataObject", "vbDate", "vbDecimal", _
		"vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", _
		"vbDefaultButton4", "vbDouble", "vbEmpty", "vbError", _
		"vbExclamation", "vbFalse", "vbFirstFourDays", "vbFirstFullWeek", _
		"vbFirstJan1", "vbFormFeed", "vbFriday", "vbGeneralDate", "vbGreen", _
		"vbIgnore", "vbInformation", "vbInteger", "vbLf", "vbLong", _
		"vbLongDate", "vbLongTime", "vbMagenta", "vbMonday", "vbNewLine", _
		"vbNo", "vbNull", "vbNullChar", "vbNullString", "vbObject", _
		"vbObjectError", "vbOK", "vbOKCancel", "vbOKOnly", "vbQuestion", _
		"vbRed", "vbRetry", "vbRetryCancel", "vbSaturday", "vbShortDate", _
		"vbShortTime", "vbSingle", "vbString", "vbSunday", "vbSystemModal", _
		"vbTab", "vbTextCompare", "vbThursday", "vbTrue", "vbTuesday", _
		"vbUseDefault", "vbUseSystem", "vbUseSystemDayOfWeek", "vbVariant", _
		"vbVerticalTab", "vbWednesday", "vbWhite", "vbYellow", "vbYes", _
		"vbYesNo", "vbYesNoCancel" _
		)
	' Define a list of Functions
	lvFunctions = Array( _
		"Abs", "Array", "Asc", "Atn", "CBool", "CByte", "CCur", "CDate", _
		"CDbl", "Chr", "CInt", "CLng", "Cos", "CreateObject", "CSng", "CStr", _
		"Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue", _
		"Day", "Exp", "Filter", "Fix", "FormatCurrency", "FormatDateTime", _
		"FormatNumber", "FormatPercent", "GetObject", "Hex", "Hour", _
		"InputBox", "InStr", "InStrRev", "Int", "IsArray", "IsDate", _
		"IsEmpty", "IsNull", "IsNumeric", "IsObject", "Join", "LBound", _
		"LCase", "Left", "Len", "LoadPicture", "Log", "LTrim", "Mid", _
		"Minute", "Month", "MonthName", "MsgBox", "Now", "Oct", "Replace", _
		"RGB", "Right", "Rnd", "Round", "RTrim", "ScriptEngine", _
		"ScriptEngineBuildVersion", "ScriptEngineMajorVersion", _
		"ScriptEngineMinorVersion", "Second", "Sgn", "Sin", "Space", "Split", _
		"Sqr", "StrComp", "StrReverse", "String", "Tan", "TimeSerial", _
		"TimeValue", "Trim", "TypeName", "UBound", "UCase", "VarType", _
		"Weekday", "WeekdayName", "Year" _
		)
	' Define a list of Methods
	lvMethods = Array( _
		"Application.Lock", "Application.UnLock", "Err.Clear", "Err.Raise", _
		"Request.BinaryRead", "Response.AddHeader", "Response.AppendToLog", _
		"Response.BinaryWrite", "Response.Clear", "Response.End", _
		"Response.Flush", "Response.IsClientConnected", "Response.Pics", _
		"Response.Redirect", "Response.Write", "ScriptEngine", _
		"ScriptEngineBuildVersion", "ScriptEngineMajorVersion", _
		"ScriptEngineMinorVersion", "Server.CreateObject", _
		"Server.HTMLEncode", "Server.MapPath", "Server.URLEncode", _
		"Server.URLPathEncode", "Session.Abandon" _
		)
	' Define a list of properties
	lvProperties = Array( _
		"Application.Contents", "Application.Contents.Count", _
		"Application.Contents.Item", "Application.Contents.Key", _
		"Application.StaticObjects", "Application.StaticObjects.Count", _
		"Application.StaticObjects.Item", "Application.StaticObjects.Key", _
		"Application.Value", "Err.description", "Err.helpcontext", _
		"Err.helpfile", "Err.number", "Err.source", _
		"Request.ClientCertificate", "Request.ClientCertificate.Count", _
		"Request.ClientCertificate.Item", "Request.ClientCertificate.Key", _
		"Request.Cookies", "Request.Cookies.Count", "Request.Cookies.Item", _
		"Request.Cookies.Key", "Request.Form", "Request.Form.Count", _
		"Request.Form.Item", "Request.Form.Key", "Request.Item", _
		"Request.QueryString", "Request.QueryString.Count", _
		"Request.QueryString.Item", "Request.QueryString.Key", _
		"Request.ServerVariables", "Request.ServerVariables.Count", _
		"Request.ServerVariables.Item", "Request.ServerVariables.Key", _
		"Request.TotalBytes", "Response.Buffer", "Response.CacheControl", _
		"Response.CharSet", "Response.ContentType", "Response.Cookies", _
		"Response.Cookies.Count", "Response.Cookies.Item", _
		"Response.Cookies.Key", "Response.Expires", _
		"Response.ExpiresAbsolute", "Response.Status", _
		"Server.ScriptTimeout", "Session.CodePage", _
		"Session.Contents", "Session.Contents.Count", _
		"Session.Contents.Item", "Session.Contents.Key", "Session.LCID", _
		"Session.SessionID", "Session.StaticObjects", _
		"Session.StaticObjects.Count", "Session.StaticObjects.Item", _
		"Session.StaticObjects.Key", "Session.Timeout", "Session.Value" _
		)
	' Define a list of operators
	lvOperators = Array("+", "=", "&", "/", "^", "Imp", "\", "Is", "Mod", _
		"*", "-", "Not", "Or", "-", "Xor", ">", "<")

	' Define a list of statements
	lvStatements = Array( _
		"Call", "Case", "Const", "Dim", "Do", "Loop", "Erase", "Exit", "For", _
		"Next", "Each", "Function", "If", "Then", "Else", "On", "Error", _
		"Option", "Explicit", "Private", "Public", "Randomize", "ReDim", _
		"Select", "Set", "Sub", "While", "Wend" _
		)

	' Split NewCode into an array
	lvLines = Split(NewCode, vbCrLf)
	
	'Loop Through Each Line
	For Each lsLine In lvLines
		lbBuildString = False     ' No, we are not within a string
		lbBuildWord = False    ' No, we are not within a word
		lsNewLine = ""      ' Reset the formated line
		lsWord = ""         ' Reset the word
		lsString = ""		' Reset string text
'		lbBuildComment = False	' No, we are not within a comment
		
		' Determine if line has text.
		If Len(lsLine) > 0 Then
			' Loop through each character in the line.
			For lnStart = 1 To Len(lsLine)
				lsChar = Mid(lsLine, lnStart, 1) ' Grab current character
				
				If lbBuildString Then ' Are we currently in a string?
					lsString = lsString & lsChar
					' Determine if this character gets us out of the string.
					If lsChar = lsSTRING_CHAR Then
						lsNewLine = lsNewLine & _
							"<SPAN class=""vbScript-String"">" & _
							Server.HTMLEncode(lsString) & _
							"</SPAN>"
						lsString = ""
						lbBuildString = False
					End If
					' Add the current character to the line we are building.
					'lsNewLine = lsNewLine & Server.HTMLEncode(lsChar)
				ElseIf lbBuildComment Then ' Are we currently in a comment?
					lsNewLine = lsNewLine & Server.HTMLEncode(lsChar)
				Else ' We are not in a string or a comment.
					If lbBuildWord Then ' Are we currently in a word?
						' Check to see if we have encountered a seperator
						If IsInArray(lsChar, lvSeperators, False) Or IsInArray(lsChar, lvOperators, False) Then
							' Determine if word is reserved.
							If IsInArray(lsWord, lvReservedWords, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Reserved", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvConstants, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Constant", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvFunctions, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Function", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvMethods, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Method", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvProperties, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Property", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvStatements, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Statement", "vbScript", lsWord, "")
							ElseIf IsInArray(lsWord, lvOperators, False) Then
								lsNewLine = lsNewLine & LookupLink("vbScript-Operator", "vbScript", lsWord, "")
							Else
								lsNewLine = lsNewLine & lsWord
							End If
							If lsChar = lsLINE_COMMENT_CHAR Then ' We are now within a comment.
								lbBuildComment = True
								lsNewLine = lsNewLine & "<SPAN class=""Comment"">" & lsLINE_COMMENT_CHAR
							Else
								If IsInArray(lsChar, lvOperators, False) Then
									lsNewLine = lsNewLine & LookupLink("vbScript-Operator", "vbScript", lsChar, "")
								Else
									lsNewLine = lsNewLine & Server.HTMLEncode(lsChar)
								End If
							End If
							' Reset word variables
							lbBuildWord = False
							lsWord = ""
						Else
							lsWord = lsWord & lsChar
						End If
					Else
						If lsChar = lsSTRING_CHAR Then
							lbBuildString = True
							lsString = lsSTRING_CHAR
							'lsNewLine = lsNewLine & lsSTRING_CHAR
						ElseIf lsChar = lsLINE_COMMENT_CHAR Then
							lbBuildComment = True
							lsNewLine = lsNewLine & "<SPAN class=""vbScript-Comment"">" & lsLINE_COMMENT_CHAR
						ElseIf IsInArray(lsChar, lvOperators, False) Then
							lsNewLine = lsNewLine & LookupLink("vbScript-Operator", "vbScript", lsChar, "")
						Else
							' Are we in a seperator?
							If IsInArray(lsChar, lvSeperators, False) Then
								lsNewLine = lsNewLine & Server.HTMLEncode(lsChar)
							Else
								lsWord = lsChar
								lbBuildWord = True
							End If
						End If
					End If
				End If
			Next
		End If
		If lbBuildComment Then
			' Determine if next line is commented
			' Get Rid of White Space
			lsLine = Replace(lsLine, " ", "")
			lsLine = Replace(lsLine, vbTab, "")
			If Not Right(lsLine, 1) = lsLINE_COMMENT_CONTINUE_CHAR Then
				lsNewLine = lsNewLine & "</SPAN>"
				lbBuildComment = False
			End If
		ElseIf lbBuildWord Then
			' Determine if word is reserved.
			If IsInArray(lsWord, lvReservedWords, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Reserved", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvConstants, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Constant", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvFunctions, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Function", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvMethods, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Method", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvProperties, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Property", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvStatements, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Statement", "vbScript", lsWord, "")
			ElseIf IsInArray(lsWord, lvOperators, False) Then
				lsNewLine = lsNewLine & LookupLink("vbScript-Operator", "vbScript", lsWord, "")
			Else
				lsNewLine = lsNewLine & lsWord
			End If
		ElseIf lbBuildString Then
			lsNewLine = lsNewLine & _
				"<SPAN class=""vbScript-String"">" & _
				Server.HTMLEncode(lsString) & _
				"</SPAN>"
			lsString = ""
			lbBuildString = False
		End If
		'lsNewLine = Replace(lsNewLine, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
		'Response.Write ". "
		lsResult = lsResult & lsNewLIne & vbCrLf
	Next
	lsResult = "<SPAN class=""vbScript""><PRE>" & lsResult & "</PRE></SPAN>"
	'Response.Write "<HR>"
	'lsResult = Replace(lsResult, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
	ColorVBScript = lsResult
	' = Join(lvLines, "<BR>")
End Function
' -----------------------------------------------------------------------------
Function LookupLink(ByRef ClassName, ByRef ScriptLanguage, ByRef ReservedWord, ByRef Argument)
	Const lsURL_LOOKUP = "find.asp"


	Dim lsResult
	lsResult = _
		"<A" & _
			" class=""" & ClassName & """" & _
			" href=""javascript://"" onclick=""lookUp('" & Server.URLEncode(ReservedWord) & "')"">"
			
	' Determine wich word to show.
	If Not Argument = "" Then
		lsResult = lsResult & Argument & "</A>"
	Else
		lsResult = lsResult & ReservedWord & "</A>"
	End If
	LookupLink = lsResult
End Function
' -----------------------------------------------------------------------------
Function IsInArray(ByRef LookupWord, ByRef ArrayList, ByRef MatchCase)
	Dim lsWord
	For Each lsWord In ArrayList
		If MatchCase And LookupWord = lsWord Then
			IsInArray = True
			Exit For
		ElseIf Not(MatchCase) And(LCase(LookupWord) = LCase(lsWord))Then
			IsInArray = True
			LookupWord = lsWord ' Make correct Case
			Exit For
		End If
	Next
End Function
' -----------------------------------------------------------------------------
%>