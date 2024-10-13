<%
Class clsVBColor
	
	Private mStrReservedAry		' Array of Reserved Words in this langauge
	Private mStrConstantAry		' Array of recognized constants
	Private mStrFunctionAry		' Array of recognized functions
	Private mStrMethodAry		' Array of recognized methods
	Private mStrPropertyAry		' Array of recognized properties
	Private mStrStatementAry	' Array of recognized statments
	Private mStrOperatorAry		' Array of recognized operators
	Private mStrSeperatorAry	' Array of single characters that seperate words.
	
	Private mStrStringCharacter
	Private mStrLineCommentCharacter
	Private mStrLineCommentContinueCharacter
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		mStrStringCharacter = """"
		mStrLineCommentCharacter = "'"
		mStrLineCommentContinueCharacter = "_"

		' Define a list of characters that seperate words from next character.
		mStrSeperatorAry = Array(" ", vbTab, ":", "(", ")", ",", "'")

		' Define a list of Reserved Words
		mStrReservedAry = Array( _
			"And", "Array", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", _
			"Case", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", _
			"Class", "Close", "CInt", "CLng", "Const", "CSng", "CStr", _
			"Currency", "CVar", "Date", "Decimal", "Declare", "DefBool", _
			"DefByte", "DefCur", "DefDate", "DefDbl", "DefInt", "DefLng", _
			"DefSng", "DefStr", "DefVar", "Dim", "Do", "Double", "Each", _
			"Else", "ElseIf", "Empty", "End", "Enum", "Eqv", "Erase", "Error", _
			"Exit", "Explicit", "False", "Fix", "For", "Friend", "Function", _
			"Get", "Global", "If", "Imp", "In", "Input", "InputB", "Int", _
			"Integer", "Is", "LBound", "Len", "LenB", "Local", "Lock", "Long", _
			"Loop", "Me", "Mod", "New", "Next", "Not", "Nothing", "Null", _
			"Open", "Option", "Optional", "Or", "On", "Public", "Print", _
			"Private", "Property", "ReDim", "Resume", "Select", "Set", "Sgn", _
			"Single", "String", "Sub", "Then", "To", "True", "Type", "UBound", _
			"UnLock", "Variant", "Wend", "While", "With", "Write", _
			"Implements", "Compare", "Text", "Let" _
			)
			
		'mStrReservedAry = Array()
		' Define a list of Constants
		mStrConstantAry = Array( _
			"vbAbort", "vbAbortRetryIgnore", "vbApplicationModal", "vbArray", _
			"vbBinaryCompare", "vbBlack", "vbBlue", "vbBoolean", "vbByte", _
			"vbCancel", "vbCr", "vbCritical", "vbCrLf", "vbCurrency", _
			"vbCyan", "vbDatabaseCompare", "vbDataObject", "vbDate", _
			"vbDecimal", "vbDefaultButton1", "vbDefaultButton2", _
			"vbDefaultButton3", "vbDefaultButton4", "vbDouble", "vbEmpty", _
			"vbError", "vbExclamation", "vbFalse", "vbFirstFourDays", _
			"vbFirstFullWeek", "vbFirstJan1", "vbFormFeed", "vbFriday", _
			"vbGeneralDate", "vbGreen", "vbIgnore", "vbInformation", _
			"vbInteger", "vbLf", "vbLong", "vbLongDate", "vbLongTime", _
			"vbMagenta", "vbMonday", "vbNewLine", "vbNo", "vbNull", _
			"vbNullChar", "vbNullString", "vbObject", "vbObjectError", "vbOK", _
			"vbOKCancel", "vbOKOnly", "vbQuestion", "vbRed", "vbRetry", _
			"vbRetryCancel", "vbSaturday", "vbShortDate", "vbShortTime", _
			"vbSingle", "vbString", "vbSunday", "vbSystemModal", "vbTab", _
			"vbTextCompare", "vbThursday", "vbTrue", "vbTuesday", _
			"vbUseDefault", "vbUseSystem", "vbUseSystemDayOfWeek", _
			"vbVariant", "vbVerticalTab", "vbWednesday", "vbWhite", _
			"vbYellow", "vbYes", "vbYesNo", "vbYesNoCancel" _
			)
		' Define a list of Functions
		mStrFunctionAry = Array( _
			"Abs", "Array", "Asc", "Atn", "CBool", "CByte", "CCur", "CDate", _
			"CDbl", "Chr", "CInt", "CLng", "Cos", "CreateObject", "CSng", _
			"CStr", "Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", _
			"DateValue", "Day", "Exp", "Filter", "Fix", "FormatCurrency", _
			"FormatDateTime", "FormatNumber", "FormatPercent", "GetObject", _
			"Hex", "Hour", "InputBox", "InStr", "InStrRev", "Int", "IsArray", _
			"IsDate", "IsEmpty", "IsNull", "IsNumeric", "IsObject", "Join", _
			"LBound", "LCase", "Left", "Len", "LoadPicture", "Log", "LTrim", _
			"Mid", "Minute", "Month", "MonthName", "MsgBox", "Now", "Oct", _
			"Replace", "RGB", "Right", "Rnd", "Round", "RTrim", _
			"ScriptEngine", "ScriptEngineBuildVersion", _
			"ScriptEngineMajorVersion", "ScriptEngineMinorVersion", "Second", _
			"Sgn", "Sin", "Space", "Split", "Sqr", "StrComp", "StrReverse", _
			"String", "Tan", "TimeSerial", "TimeValue", "Trim", "TypeName", _
			"UBound", "UCase", "VarType", "Weekday", "WeekdayName", "Year" _
			)
		' Define a list of Methods
		mStrMethodAry = Array( _
			"Application.Lock", _
			"Application.UnLock", _
			"Err.Clear", _
			"Err.Raise", _
			"Request.BinaryRead", _
			"Response.AddHeader", _
			"Response.AppendToLog", _
			"Response.BinaryWrite", _
			"Response.Clear", _
			"Response.End", _
			"Response.Flush", _
			"Response.IsClientConnected", _
			"Response.Pics", _
			"Response.Redirect", _
			"Response.Write", _
			"ScriptEngine", _
			"ScriptEngineBuildVersion", _
			"ScriptEngineMajorVersion", _
			"ScriptEngineMinorVersion", _
			"Server.CreateObject", _
			"Server.HTMLEncode", _
			"Server.MapPath", _
			"Server.URLEncode", _
			"Server.URLPathEncode", _
			"Session.Abandon" _
			)
		' Define a list of properties
		mStrPropertyAry = Array( _
			"Application.Contents", _
			"Application.Contents.Count", _
			"Application.Contents.Item", _
			"Application.Contents.Key", _
			"Application.StaticObjects", _
			"Application.StaticObjects.Count", _
			"Application.StaticObjects.Item", _
			"Application.StaticObjects.Key", _
			"Application.Value", _
			"Err.description", _
			"Err.helpcontext", _
			"Err.helpfile", _
			"Err.number", _
			"Err.source", _
			"Request.ClientCertificate", _
			"Request.ClientCertificate.Count", _
			"Request.ClientCertificate.Item", _
			"Request.ClientCertificate.Key", _
			"Request.Cookies", _
			"Request.Cookies.Count", _
			"Request.Cookies.Item", _
			"Request.Cookies.Key", _
			"Request.Form", _
			"Request.Form.Count", _
			"Request.Form.Item", _
			"Request.Form.Key", _
			"Request.Item", _
			"Request.QueryString", _
			"Request.QueryString.Count", _
			"Request.QueryString.Item", _
			"Request.QueryString.Key", _
			"Request.ServerVariables", _
			"Request.ServerVariables.Count", _
			"Request.ServerVariables.Item", _
			"Request.ServerVariables.Key", _
			"Request.TotalBytes", _
			"Response.Buffer", _
			"Response.CacheControl", _
			"Response.CharSet", _
			"Response.ContentType", _
			"Response.Cookies", _
			"Response.Cookies.Count", _
			"Response.Cookies.Item", _
			"Response.Cookies.Key", _
			"Response.Expires", _
			"Response.ExpiresAbsolute", _
			"Response.Status", _
			"Server.ScriptTimeout", _
			"Session.CodePage", _
			"Session.Contents", _
			"Session.Contents.Count", _
			"Session.Contents.Item", _
			"Session.Contents.Key", _
			"Session.LCID", _
			"Session.SessionID", _
			"Session.StaticObjects", _
			"Session.StaticObjects.Count", _
			"Session.StaticObjects.Item", _
			"Session.StaticObjects.Key", _
			"Session.Timeout", _
			"Session.Value" _
			)
		' Define a list of operators
		mStrOperatorAry = Array("+", "=", "&", "/", "^", "Imp", "\", "Is", "Mod", _
			"*", "-", "Not", "Or", "-", "Xor", ">", "<")

		' Define a list of statements
		mStrStatementAry = Array( _
			"Call", "Case", "Const", "Dim", "Do", "Loop", "Erase", "Exit", _
			"For", "Next", "Each", "Function", "If", "Then", "Else", "On", _
			"Error", "Option", "Explicit", "Private", "Public", "Randomize", _
			"ReDim", "Select", "Set", "Sub", "While", "Wend" _
			)
	End Sub
' -----------------------------------------------------------------------------
	Public Function Colorize(ByRef NewCode)

		
		Dim lStrLine			' A single line within the code passed to the procedure
		Dim lStrCharacter		' A single character within the current line being looked at.
		Dim lStrString			' Contents of a string
		Dim lLngStart			' Character Start Position that we are looking at.
		Dim lStrWord			' A word being formed from the character being looked at.
		Dim lBlnBuildString		' Are we building a string?
		Dim lBlnBuildWord		' Are we building a word?
		Dim lStrNew			' The new line that is being build from a line sent to this procedure.
		Dim lBlnBuildComment	' Are we building a comment?
		Dim lStrResult			' The results to return when this procedure is finnished.
		Dim lLngRealPosition
		Dim lStrLines			' Array of lines sent to this procedure

		' Split NewCode into an array
		lStrLines = Split(NewCode, vbCrLf)
		
		'Loop Through Each Line
		For Each lStrLine In lStrLines
			lBlnBuildString = False		' No, we are not within a string
			lBlnBuildWord = False		' No, we are not within a word
			lStrNew = ""				' Reset the formated line
			lStrWord = ""				' Reset the word
			lStrString = ""				' Reset string text
'			lBlnBuildComment = False	' No, we are not within a comment
			
			' Determine if line has text.
			If Len(lStrLine) > 0 Then
				' Loop through each character in the line.
				lLngRealPosition = 0
				For lLngStart = 1 To Len(lStrLine)
					lLngRealPosition = lLngRealPosition + 1
					lStrCharacter = Mid(lStrLine, lLngStart, 1) ' Grab current character
					
					If lBlnBuildString Then ' Are we currently in a string?
						lStrString = lStrString & lStrCharacter
						' Determine if this character gets us out of the string.
						If lStrCharacter = mStrStringCharacter Then
							lStrNew = lStrNew & _
								"<SPAN class=""String"">" & _
								Server.HTMLEncode(lStrString) & _
								"</SPAN>"
							lStrString = ""
							lBlnBuildString = False
						End If
						' Add the current character to the line we are building.
						'lStrNew = lStrNew & Server.HTMLEncode(lStrCharacter)
					ElseIf lBlnBuildComment Then ' Are we currently in a comment?
						lStrNew = lStrNew & Server.HTMLEncode(lStrCharacter)
					Else ' We are not in a string or a comment.
						If lBlnBuildWord Then ' Are we currently in a word?
							' Check to see if we have encountered a seperator
							If IsInArray(lStrCharacter, mStrSeperatorAry, False) _
							Or IsInArray(lStrCharacter, mStrOperatorAry, False) Then
								' Determine if word is reserved.
								lStrNew = lStrNew & Word(lStrWord)
								If lStrCharacter = mStrLineCommentCharacter Then 
									' We are now within a comment.
									lBlnBuildComment = True
									lStrNew = lStrNew & "<SPAN class=""Comment"">" & _
										mStrLineCommentCharacter
								ElseIf IsInArray(lStrCharacter, mStrOperatorAry, False) Then
									lStrNew = lStrNew & "<SPAN class=""Operator"">" & _
										Server.HTMLEncode(lStrCharacter) & "</SPAN>"
								ElseIf lStrCharacter = vbTab Then
									lStrNew = lStrNew & Tabs(lLngRealPosition)
								Else
									lStrNew = lStrNew & Server.HTMLEncode(lStrCharacter)
								End If
								' Reset word variables
								lBlnBuildWord = False
								lStrWord = ""
							Else
								lStrWord = lStrWord & lStrCharacter
							End If
						Else
							If lStrCharacter = mStrStringCharacter Then
								lBlnBuildString = True
								lStrString = mStrStringCharacter
								'lStrNew = lStrNew & mStrStringCharacter
							ElseIf lStrCharacter = mStrLineCommentCharacter Then
								lBlnBuildComment = True
								lStrNew = lStrNew & "<SPAN class=""Comment"">" & _
									mStrLineCommentCharacter
							ElseIf IsInArray(lStrCharacter, mStrOperatorAry, False) Then
								lStrNew = lStrNew & "<SPAN class=""Operator"">" & _
									Server.HTMLEncode(lStrCharacter) & "</SPAN>"
							ElseIf IsInArray(lStrCharacter, mStrSeperatorAry, False) Then
								' we are in a seperator?
								If lStrCharacter = vbTab Then
									lStrNew = lStrNew & Tabs(lLngRealPosition)
								Else
									lStrNew = lStrNew & Server.HTMLEncode(lStrCharacter)
								End If
							Else
								lStrWord = lStrCharacter
								lBlnBuildWord = True
							End If
						End If
					End If
				Next
			End If
			If lBlnBuildComment Then
				' Determine if next line is commented
				' Get Rid of White Space
				lStrLine = Replace(lStrLine, " ", "")
				lStrLine = Replace(lStrLine, vbTab, "")
				If Not Right(lStrLine, 1) = mStrLineCommentContinueCharacter Then
					lStrNew = lStrNew & "</SPAN>"
					lBlnBuildComment = False
				End If
			ElseIf lBlnBuildWord Then
				' Determine if word is reserved.
				lStrNew = lStrNew & Word(lStrWord)
			ElseIf lBlnBuildString Then
				lStrNew = lStrNew & _
					"<SPAN class=""String"">" & _
					Server.HTMLEncode(lStrString) & _
					"</SPAN>"
				lStrString = ""
				lBlnBuildString = False
			End If
			'lStrNew = Replace(lStrNew, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
			'Response.Write ". "
			lStrResult = lStrResult & lStrNew & "<BR>" & vbCrLf
		Next
		lStrResult = "<DIV class=""vbScript"">" & lStrResult & "</DIV>"
		'Response.Write "<HR>"
		lStrResult = Replace(lStrResult, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
		Colorize = lStrResult
		' = Join(lStrLines, "<BR>")
	End Function
' -----------------------------------------------------------------------------
	Private Function IsInArray(ByRef pStrLookupWord, ByRef pStrListAry, _
		ByRef pBlnMatchCase)

		Dim lStrWord
		For Each lStrWord In pStrListAry
			If pBlnMatchCase And pStrLookupWord = lStrWord Then
				IsInArray = True
				Exit For
			ElseIf Not(pBlnMatchCase) And(LCase(pStrLookupWord) = LCase(lStrWord))Then
				IsInArray = True
				pStrLookupWord = lStrWord ' Make correct Case
				Exit For
			End If
		Next

	End Function
' -----------------------------------------------------------------------------
	Private Function Tabs(ByRef pLngRealPosition)
		Dim lLngIndex
		Dim lLngSpaces
		' Hard-coded hack.  Need to do research for dynamic tab spacing.
		Select Case pLngRealPosition Mod 4
			Case 0
				pLngRealPosition = pLngRealPosition + 0
				lLngSpaces = 1
			Case 1
				pLngRealPosition = pLngRealPosition + 3
				lLngSpaces = 4
			Case 2
				pLngRealPosition = pLngRealPosition + 2
				lLngSpaces = 3
			Case 3
				pLngRealPosition = pLngRealPosition + 1
				lLngSpaces = 2
		End Select
		For lLngIndex = 1 to lLngSpaces
			Tabs = Tabs & "&nbsp;"
		Next
	End Function
' -----------------------------------------------------------------------------
	Private Function Word(ByRef pStrWord)
		Dim lStrWordType
		If IsInArray(pStrWord, mStrReservedAry, False) Then
			lStrWordType = "Reserved"
		ElseIf IsInArray(pStrWord, mStrConstantAry, False) Then
			lStrWordType = "Constant"
		ElseIf IsInArray(pStrWord, mStrFunctionAry, False) Then
			lStrWordType = "Function"
		ElseIf IsInArray(pStrWord, mStrMethodAry, False) Then
			lStrWordType = "Method"
		ElseIf IsInArray(pStrWord, mStrPropertyAry, False) Then
			lStrWordType = "Property"
		ElseIf IsInArray(pStrWord, mStrStatementAry, False) Then
			lStrWordType = "Statement"
		ElseIf IsInArray(pStrWord, mStrOperatorAry, False) Then
			lStrWordType = "Operator"
		Else
			lStrWordType = ""
		End If
		Word = "<SPAN class=""" & lStrWordType & """>" & Server.HTMLEncode(pStrWord) & "</SPAN>"
	End Function
End Class
' -----------------------------------------------------------------------------
%>