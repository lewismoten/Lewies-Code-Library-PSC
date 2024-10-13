<!--METADATA
	TYPE="TypeLib"
	NAME="Microsoft ActiveX Data Objects 2.5 Library"
	UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
	VERSION="2.5"
-->
<%
' -----------------------------------------------------------------------------
	Class clsDatabase
' -----------------------------------------------------------------------------

		' Define Read/Write properties
		Dim ParseMetaData				' Determines if field names/types are parsed

		' Define internal variables
		Private mBlnIsSQLServer			' determines if we are talking to an SQL server

		Private mLngAbsolutePage		' determines current position of absolute page
		Private mLngFieldTypesAry		' Array of Field Types
		Private mLngPageCount			' Number of pages returned by recordset
		Private mLngPageSize			' determines amount of records are returned on each page

		Private mObjConn				' Database connection object
		Private mObjRs					' Database Recordset

		Private mStrConnectionString		' Connection string to database
		Private mStrErrors				' Collection of internal errors
		Private mStrFieldNamesAry		' Array of Field Names
		Private mStrPassword			' Password to access database
		Private mLngRecordsAffected		' Amount of records affected by SQL Query
		Private mStrUserID				' UserID to login to database with
		Private mStrVersion				' Version of the code
		
' -----------------------------------------------------------------------------
		Public Property Get ClassName()
			ClassName = "LewisMoten.clsDatabase"
		End Property
' -----------------------------------------------------------------------------
		Private Sub Class_Initialize()
			
			' Initialize exposed properties
			ParseMetaData = False

			' Initialize internal properties
			mStrFieldNamesAry		= Array()
			mLngFieldTypesAry		= Array()
			mLngRecordsAffected		= -1
			mLngAbsolutePage		= 1
			mLngPageSize			= -1
			mLngPageCount			= -1
			mStrVersion				= 3.0
			
			' Grab database connection information
			mStrConnectionString	= _
				"DRIVER=Microsoft Access Driver (*.mdb);" & _
				"UserCommitSync=Yes;" & _
				"Threads=3;" & _
				"SafeTransactions=0;" & _
				"PageTimeout=5;" & _
				"MaxScanRows=8;" & _
				"MaxBufferSize=2048;" & _
				"UID=" & lstrDBUserName & ";" & _
				"PWD=" & lstrDBPassword & ";" & _
				"DefaultDir=" & Server.MapPath("/Tests/PickFolder/") & ";" & _
				"DBQ=" & Server.MapPath("/Tests/PickFolder/Folders.mdb")

			mStrUserID				= "Admin"
			mStrPassword			= ""
			mBlnIsSQLServer			= False
			
			' Attempt to open database connection
			Call OpenDatabase()
			
		End Sub ' Class_Initialize()
' -----------------------------------------------------------------------------
		Private Sub Class_Terminate()
			
			' Close the database connection
			CloseDatabase()
		
		End Sub ' Class_Terminate()
' -----------------------------------------------------------------------------
		Private Sub DBErr(ByVal pstrErr)
		
			pstrErr = Replace(pstrErr, "[Microsoft][ODBC Microsoft Access Driver]", "")
		
			Response.Clear
			Response.Write "<FONT color=""Red"">Oops!</FONT><BR>"
			Response.Write "<P>It appears an error occured within our database.</P>"
			Response.Write "<BLOCKQUOTE><CODE>" & pstrErr & "</CODE></BLOCKQUOTE>"
			Response.End
		End Sub
		Public Property Get AbsolutePage()
			
			AbsolutePage = mLngAbsolutePage
		
		End Property ' Get AbsolutePage()
' -----------------------------------------------------------------------------
		Public Property Let AbsolutePage(ByRef anAbsolutePage)
			
			mLngAbsolutePage = anAbsolutePage
		
		End Property ' Let AbsolutePage()
' -----------------------------------------------------------------------------
		Public Sub CloseDatabase()
			
			' Prepare for errors
			On Error Resume Next
			
			' If the recordset object is defined
			If IsObject(mObjRs) Then
					
				' If the recordset object is open
				If mObjRs.State = ADODB.adStateOpen Then
						
					' Close Recordset
					mObjRs.Close
					
				End If ' mObjRs.State = ADODB.adStateOpen
					
				' Release object from memory
				Set mObjRs = Nothing

			End If ' IsObject(mObjRs)
			
			' If connection object has been created
			If IsObject(mObjConn) Then
				
				' If the connection object is open
				If mObjConn.State = ADODB.adStateOpen Then
					
					' Close the connection
					mObjConn.Close
				
				End If ' mObjConn.State = ADODB.adStateOpen
				
				' Release object from memory
				Set mObjConn = Nothing
				
			End If ' IsObject(mObjConn)
			
		End Sub ' CloseDatabase()
' -----------------------------------------------------------------------------
		Public Property Get ConcatenationOperator()
			If mBlnIsSQLServer Then
				ConcatenationOperator = "+"
			Else
				ConcatenationOperator = "&"
			End If
		End Property
' -----------------------------------------------------------------------------
		Public Property Get ConnectionString()

			ConnectionString = mStrConnectionString
		
		End Property ' Get ConnectionString()
' -----------------------------------------------------------------------------
		Public Property Let ConnectionString(ByRef asConnectionString)
		
			mStrConnectionString = ConnectionString
		
		End Property ' Let ConnectionString(ByRef asConnectionString)
' -----------------------------------------------------------------------------
		Public Property Get Errors()

			' If no errors are defined
			If Len(mStrErrors) <= 1 Then
			
				' Do nothing
			
			' Else errors are present
			Else
			
				' If errors begin with a delimiter
				If Left(mStrErrors, 1) = ";" Then
					
					' Return the errors with the first character removed
					Errors = Mid(mStrErrors, 2)
			
				' Else errors do not begin with a delimiter
				Else
					
					' Return errors
					Errors = mStrErrors
				
				End If ' Left(mStrErrors, 1) = ";"
				
			End If ' Len(mStrErrors) <= 1
			
		End Property ' Get Errors()
' -----------------------------------------------------------------------------
		Private Function SQLDebugger(pStrFunction, pStrSQL)
			'DebugThis "clsDatabase." & pStrFunction, pStrSQL
		End Function
' -----------------------------------------------------------------------------
		Public Function ExecuteSQL(ByRef asSQL)
			
			SQLDebugger "ExecuteSQL()", asSQL
			
			' Setup default values (in case procedure failes)
			mLngRecordsAffected = -1
			ExecuteSQL = False

			' Prepare for errors when executing query
			On Error Resume Next

			Call mObjConn.Execute(asSQL, mLngRecordsAffected, ADODB.adCmdText Or ADODB.adExecuteNoRecords)
		
			' If errors occured
			If Err Then

				' Append error to internal error collection
				Call DBErr(Err.Description)

				' Clear the error
				Err.Clear
			
			' Else SQL query executed fine
			Else
				
				' Return positive results
				ExecuteSQL = True
				
			End If ' Err
		
		End Function ' ExecuteSQL(ByRef asSQL)
' -----------------------------------------------------------------------------
		Public Function FieldNames()
		
			FieldNames = mStrFieldNamesAry
			
		End Function ' FieldNames()
' -----------------------------------------------------------------------------
		Public Function FieldTypes()
			
			Dim lnIndex		' Index of Field Types Array
			Dim lnValue		' Value of number within field types array
			
			' Loop through each index in the field types array
			For lnIndex = 0 To UBound(mLngFieldTypesAry, 1)
			
				' Grab the value of the current index
				lnValue = mLngFieldTypesAry(lnIndex)
				
				' If the value is a number
				If IsNumeric(lnValue) And Not lnValue = "" Then
					
					' Convert SubType String to a SubType Long
					mLngFieldTypesAry(lnIndex) = CLng(lnValue)
				
				End If ' IsNumeric(lnValue) And Not lnValue = ""
				
			Next ' lnIndex
			
			' Return the values of the field types
			FieldTypes = mLngFieldTypesAry
			
		End Function ' FieldTypes()
' -----------------------------------------------------------------------------
		Public Function FormatBoolean(ByVal abExpression, ByRef abDefault)

			Dim lbTrue
			Dim lbFalse
			
			' If database is an SQL Server
			If mBlnIsSQLServer Then
			
				' Assign value for true
				lbTrue = 1
			
			' Else database is an Access database
			Else
				
				' Assign value for true
				lbTrue = -1
			
			End If ' mBlnIsSQLServer
			
			' Assign value for false
			lbFalse = 0

			Dim lbResult ' Result to be passed back

			' Prepare for any errors that may occur
			On Error Resume Next

			' If expression not provided
			If abExpression = "" Then

				' Set expression to default value
				abExpression = abDefault

			End If ' abExpression = ""

			' Attempt to convert expression
			lbResult = CBool(abExpression)

			' If Err Occured
			If Err Then

				' Clear the error
				Err.Clear

				' Determine action based on Expression
				Select Case LCase(abExpression)
					' True expressions
					Case "yes", "on", "true", "-1", "1"
						lbResult = True
					' False expressions
					Case "no", "off", "false", "0"
						lbResult = False
					' Unknown expression
					Case Else
						lbResult = abDefault
				End Select ' LCase(abExpression)

			End If ' Err

			' If result is True
			If lbResult Then

				' Return True
				FormatBoolean = lbTRUE

			' Else Result is false
			Else

				' Return False
				FormatBoolean = lbFALSE

			End If ' lbResult

		End Function ' FormatBoolean(ByVal abExpression, ByRef abDefault)
' -----------------------------------------------------------------------------
		Public Function FormatDate(ByRef adExpression)

			' If Expression valid date
			If IsDate(adExpression) Then

				' If database is an SQL Server
				If mBlnIsSQLServer Then
				
					' Return formated date for SQL Server
					FormatDate = "'" & adExpression & "'"
				
				' Else database is an Access File
				Else
					
					' Return formated date for Access File
					FormatDate = "#" & adExpression & "#"
				
				End If ' mBlnIsSQLServer

			' Else Expression not valid date
			Else

				' Return NULL
				FormatDate = "NULL"

			End If ' IsDate(adExpression)

		End Function ' FormatDate(ByRef adExpression)
' -----------------------------------------------------------------------------
		Public Function FormatMoney(ByVal acExpression)

			' If Empty Expression
			If acExpression = "" Then

				' Return Null
				FormatMoney = "NULL"

			' Else expression has content
			Else

				' Prepare for Errors
				On Error Resume Next

				' Attempt to convert expression to currency
				FormatMoney = CCur(acExpression)

				' If error occured
				If Err Then

					' Clear Error
					Err.Clear

					FormatMoney = "NULL"

				End If ' Err

			End If ' acExpression = ""

		End Function ' FormatMoney(ByVal acExpression)
' -----------------------------------------------------------------------------
		Public Function FormatNumeric(ByVal anExpression)

			' If expression numeric
			If IsNumeric(anExpression) And Not anExpression = "" Then

				' Return number
				FormatNumeric = anExpression

			' Else expression not numeric
			Else

				' Return NULL
				FormatNumeric = "NULL"

			End If ' IsNumeric(anExpression) And Not anExpression = ""

		End Function ' FormatNumeric(ByVal anExpression)
' -----------------------------------------------------------------------------
		Public Function FormatString(ByVal asExpression, ByRef anMaxLength)

			' If maximum length is defined
			If anMaxLength > 0 Then

				' Trim expression to maximum length
				asExpression = Left(asExpression, anMaxLength)

			End If ' anMaxLength > 0

			' Double quote SQL quote characters
			asExpression = Replace(asExpression, "'", "''")

			' If Expression is Empty
			If asExpression = "" Then

				' Return a NULL value
				FormatString = "NULL"

			' Else expression is not empty
			Else

				' Return quoted expression
				FormatString = "'" & asExpression & "'"

			End If ' asExpression

		End Function ' FormatString(ByVal asExpression, ByRef anMaxLength)
' -----------------------------------------------------------------------------
		Public Function OpenDatabase()

			' Setup default value as false
			OpenDatabase = False
			
			' If a connection string has not been defined
			If mStrConnectionString = "" Then
				
				' Append error to error collection
				Call DBErr("No connection string has been defined.")
			
				SQLDebugger "OpenDatabase()", "No connection string has been defined."
			
				' Exit the procedure
				Exit Function
				
			End If ' mStrConnectionString = ""

			SQLDebugger "OpenDatabase()", "Connection String: " & mStrConnectionString
			
			' If a previouse connection object was created
			If IsObject(mObjConn) Then
				
				' If an open database connection exists
				If mObjConn.State = ADODB.adStateOpen Then
					
					' Close current database connection
					mObjConn.Close
				
				End If ' mObjConn.State = ADODB.adStateOpen
			
			' Else no connection object has been created yet
			Else
			
				' Create Connection Object
				Set mObjConn = Server.CreateObject("ADODB.Connection")
				
			End If ' IsObject(mObjConn)

			' Prepare for errors
			On Error Resume Next
			
			' Open database connection
			Call mObjConn.Open(mStrConnectionString, mStrUserID, mStrPassword)

			' If any errors have occured
			If Err Then
				
				If Err Then
				
					SQLDebugger "OpenDatabase()", Err.Description
					
					' Append error to internal error collection
					Call DBErr(Err.Description)
					
					' Clear errors
					Err.Clear
				End If
				
				' Release connection object
				Set mObjConn = Nothing

			' Else no errrors occured.
			Else
			
				' If the recordset has not been opened
				If Not IsObject(mObjRs) Then
				
					' Create a recordset object
					Set mObjRs = Server.CreateObject("ADODB.Recordset")
				
				End If ' Not IsObject(mObjRs)
				
				' Return Positive Results
				OpenDatabase = True
				
			End If ' Err

		End Function ' OpenDatabase()
' -----------------------------------------------------------------------------
		Public Property Get PageCount()
			
			PageCount = mLngPageCount
		
		End Property ' Get PageCount()
' -----------------------------------------------------------------------------
		Public Property Get PageSize()
			
			PageSize = mLngPageSize
		
		End Property ' Get PageSize()
' -----------------------------------------------------------------------------
		Public Property Let PageSize(ByRef anPageSize)
			
			mLngPageSize = anPageSize
		
		End Property
' -----------------------------------------------------------------------------
		Public Property Get Password()
			
			Password = mStrPassword
		
		End Property ' Get Password()
' -----------------------------------------------------------------------------
		Public Property Let Password(ByRef asPassword)
			
			mStrPassword = Password
		
		End Property ' Let Password(ByRef asPassword)
' -----------------------------------------------------------------------------
		Public Function QueryFields(ByRef asFieldAry, ByVal asQuery)

			Dim loRegExp			' Regular Expression Object (requires vbScript 5.0)
			Dim loRequiredWords		' Words that MUST match within a search
			Dim loUnwantedWords		' Words that MUST NOT match within a search
			Dim loOptionalWords		' Words that AT LEAST ONE must match
			Dim lsSQL				' Arguments of SQL query that is returned (WHERE __Arguments___)
			Dim lnIndex				' Index of an array
			Dim lsPhrase			' Keyword or Phrase being worked with
			Dim lsExpressionPhrase
			Dim lsExpressionWord
			
			If Not IsArray(asFieldAry) Then Exit Function
			If Join(asFieldAry, "") = "" Then Exit Function
			
			' Create regular Expression
			Set loRegExp = New RegExp

			' Match more then once
			loRegExp.Global = True
			
			' Every letter is created equal (uppercase-lowercase = same)
			loRegExp.IgnoreCase = True
			
			lsExpressionPhrase = "[\(\[\{""].*[\)\]\}""]"
			lsExpressionWord = "[^ ]+"

			' pull out keywords and phrases that MUST match within a search
			loRegExp.Pattern = "((AND|[+&])\s*" & lsExpressionPhrase & ")|((AND\s|[+&])\s*" & lsExpressionWord & ")"
			Set loRequiredWords = loRegExp.Execute(asQuery)
			asQuery = loRegExp.Replace(asQuery, "")
			
			' pull out keywords and phrases that MUST NOT match within a search
			loRegExp.Pattern = "(((NOT|[-])\s*)?" & lsExpressionPhrase & ")|(((NOT\s+|[-])\s*)" & lsExpressionWord & ")"
			Set loUnwantedWords = loRegExp.Execute(asQuery)
			asQuery = loRegExp.Replace(asQuery, "")

			' pull out keywords and phrases that must have AT LEAST ONE match within a search
			loRegExp.Pattern = "(((OR|[|])\s*)?" & lsExpressionPhrase & ")|(((OR\s+|[|])\s*)?" & lsExpressionWord & ")"
			Set loOptionalWords = loRegExp.Execute(asQuery)
			asQuery = loRegExp.Replace(asQuery, "")
			

			' If at least 1 required word was found
			If Not loRequiredWords.Count = 0 Then
			
				' REQUIRED
				
				' Open a new group
				lsSQL = lsSQL & "("
				
				' loop through each keyword/phrase
				For lnIndex = 0 To loRequiredWords.Count - 1

					' Pull the keyword out
					lsPhrase = loRequiredWords.Item(lnIndex).Value

					' Strip boolean language
					loRegExp.Pattern = "^(AND|[+&])\s*"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					loRegExp.Pattern = "[()""\[\]{}]"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					
					' Double Quote Keyword
					lsPhrase = Replace(lsPhrase, "'", "''")

					' If we are not working with the first keyword
					If Not lnIndex = 0 Then
						
						' append logic before the keyword
						lsSQL = lsSQL & " AND "
				 	
				 	End If ' Not lnIndex = 0
				 	
				 	' Append SQL to search for the keyword within all searchable fields
					lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsPhrase & "%' OR ") & " LIKE '%" & lsPhrase & "%')"

				Next ' lnIndex
				
				' Close the group
				lsSQL = lsSQL & ")"

			End If ' Not loRequiredWords.Count = 0


			' If at least 1 optional word was found
			If Not loOptionalWords.Count = 0 Then

				' OPTIONAL

				' If the SQL query is not yet defined
				If lsSQL = "" Then
					
					' Open a new group
					lsSQL = "("
				
				' Else SQL query has content
				Else
					
					' Append logic before the group
					lsSQL = lsSQL & " AND ("
					
				End If ' lsSQL = ""

				' loop through each keyword/phrase
				For lnIndex = 0 To loOptionalWords.Count - 1

					' Pull the keyword out
					lsPhrase = loOptionalWords.Item(lnIndex).Value

					' Strip Boolean Language
					loRegExp.Pattern = "^(OR|[|])\s*"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					loRegExp.Pattern = "[()""\[\]{}]"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					
					' Double Quote the keyword
					lsPhrase = Replace(lsPhrase, "'", "''")

					
					' If we are not working with the first keyword
					If Not lnIndex = 0 Then
						
						' Append Logic before the keyword search
						lsSQL = lsSQL & " OR "
						
					End If ' Not lnIndex = 0
					
					' Append SQL to search for the keyword within all searchable fields
					lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsPhrase & "%' OR ") & " LIKE '%" & lsPhrase & "%')"

				Next ' lnIndex
				
				' Close the group
				lsSQL = lsSQL & ")"
			
			End If ' Not loOptionalWords.Count = 0

			' If at least 1 Unwanted word was found
			If Not loUnwantedWords.Count = 0 Then

				' UNWANTED

				' If the SQL query is not yet defined
				If lsSQL = "" Then
					
					' Open a new group
					lsSQL = "("
				
				' Else SQL query has content
				Else
					
					' Append logic before the group
					lsSQL = lsSQL & " AND NOT ("
					
				End If ' lsSQL = ""

				' loop through each keyword/phrase
				For lnIndex = 0 To loUnwantedWords.Count - 1

					' Pull the keyword out
					lsPhrase = loUnWantedWords.Item(lnIndex).Value

					' Strip Boolean Language
					loRegExp.Pattern = "^(NOT|[-])\s*"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					loRegExp.Pattern = "[()""\[\]{}]"
					lsPhrase = loRegExp.Replace(lsPhrase, "")
					
					' Double Quote the keyword
					lsPhrase = Replace(lsPhrase, "'", "''")

					' If we are not working with the first keyword
					If Not lnIndex = 0 Then

						' Append Logic before the keyword search
						lsSQL = lsSQL & " OR "

					End If ' Not lnIndex = 0
					
					' Append SQL to search for the keyword within all searchable fields
					lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsPhrase & "%' OR ") & " LIKE '%" & lsPhrase & "%')"

				Next ' lnIndex
				
				' Close the group
				lsSQL = lsSQL & ")"

			End If ' Not loUnwantedWords.Count = 0

			' If arguments were created
			If Not lsSQL = "" Then
				
				' Encapsilate Arguments as a group
				' in case other aguments are to be appended
				lsSQL = "(" & lsSQL & ")"
			
			End If ' Not lsSQL = ""
			
			' Return the results
			QueryFields = lsSQL
			
			' Remove object from memory
			Set loRegExp = Nothing
		
		End Function ' SearchFieldsSQL(ByRef asFieldAry, ByVal asQuery)
' -----------------------------------------------------------------------------
		Public Property Get RecordsAffected()
			
			' Return the amount of records affected by last SQL query
			RecordsAffected = mLngRecordsAffected
			
		End Property ' Get RecordsAffected()
' -----------------------------------------------------------------------------
		Public Function SetData(ByRef asSQL, ByRef avDataAry)
			
			On Error Resume Next

			Dim lnAbsolutePage		' Page number of record sets
			Dim lnPageSize			' Number of records on each page
			Dim lnPage				' Current page being worked with

			Dim lsFieldNames		' Comma-Space delimited list of field names
			Dim lsFieldTypes		' Comma-Space delmited list of field types

			' Setup default values (in case procedure failes)
			SetData				= False

			mLngFieldTypesAry		= Array()
			mLngPageCount			= -1
			mLngRecordsAffected	= -1

			mStrFieldNamesAry		= Array()
			
			' copy module variables to local ones
			lnAbsolutePage		= mLngAbsolutePage
			lnPageSize			= mLngPageSize
			
			' Reset module variables
			mLngAbsolutePage		= 1
			mLngPageSize			= -1

			If asSQL = "" Then Call("No SQL Provided!!!")
			If Err Then Call DBErr("Pre-Existing Errors: " & Err.Description)

			' Prepare for errors when executing query
			'On Error Resume Next

			' If the page size has been defined
			If lnPageSize	= -1 Then
				
				' we are to retrieve all records

				' Acquire data with ConnectionString
				
				Call mObjRs.Open(asSQL, mObjConn, ADODB.adOpenStatic, ADODB.adLockReadOnly)
				If Err Then Call DBErr(Err.Description)
				
			Else
				' Specify the Microsoft Client cursor
				mObjRs.CursorLocation = ADODB.adUseClient
				If Err Then Call DBErr(Err.Description)

				' Acquire data with ConnectionString
				Call mObjRs.Open(asSQL, mObjConn, ADODB.adOpenKeyset, ADODB.adLockOptimistic)
				If Err Then Call DBErr(Err.Description)
		
			End If ' Not lnPageSize = -1
			
			If Not mObjRs.State = ADODB.adStateClosed Then
				
				' Grab the amount of records affected
				mLngRecordsAffected = mObjRs.RecordCount
				If Err Then Call DBErr(Err.Description)
				
			End If
			
			If Err Then Call DBErr(Err.Description)
			
			' Else if the recordset is not open
			If Not mObjRs.State = ADODB.adStateOpen Then
				
				' Append error to internal error collection
				Call DBErr("Recordset could not be opened.")
				
				' Exit the procedure
				Exit Function
				
			End If ' Err
			
			' Return positive results
			SetData = True
			
			' If user wants to parse meta data
			If ParseMetaData Then
				
				' Loop through each field in the recordset
				For Each loField In mObjRs.Fields
					
					' Append the current field information to Meta-Data strings
					lsFieldNames = lsFieldNames & ", " & loField.Name
					lsFieldTypes = lsFieldTypes & ", " & loField.Type
					
				Next ' loField
				
				' Remove first comma-space delimiter
				lsFieldNames = Mid(lsFieldNames, 3)
				lsFieldTypes = Mid(lsFieldTypes, 3)
				
				' Split Meta Data and store within results
				mStrFieldNamesAry = Split(lsFieldNames, ", ")
				mLngFieldTypesAry = Split(lsFieldTypes, ", ")

			End If ' ParseMetaData
			
			' If data was found
			If Not mObjRs.EOF Then

				' If the page size has not been defined
				If lnPageSize = -1 Then
					
					mLngPageCount = 1

					' Pull data into an array
					avDataAry = mObjRs.GetRows

				' Else the page size has been defined
				Else
					
					' Define the page size
					mObjRs.PageSize = lnPageSize
					
					' Acquire the number of pages available
					mLngPageCount = mObjRs.PageCount
					
					' Jump to the page that user wants to retrieve
					mObjRs.AbsolutePage = lnAbsolutePage

					' If records still exist
					If Not mObjRs.EOF Then
						
						' Grab records from the current page
						avDataAry = mObjRs.GetRows(lnPageSize, ADODB.adBookmarkCurrent)

					End If ' Not mObjRs.EOF
					
				End If ' lnPageSize = -1

			End If ' Not mObjRs.EOF

			' If the recordset is open
			If mObjRs.State = ADODB.adStateOpen Then
			
				' Close Recordset
				mObjRs.Close
				
			End If ' mObjRs.State = ADODB.adStateOpen

		End Function ' SetData(ByRef asSQL, ByRef avDataAry)
' -----------------------------------------------------------------------------
		Public Function SetOptions(ByRef asSQL, ByRef asOptions, ByRef abEncode)

			SQLDebugger "SetOptions()", asSQL
			
			Dim lnAbsolutePage		' Page number of record sets
			Dim lnPageSize			' Number of records on each page
			Dim lnPage				' Current page being worked with

			Dim lsFieldNames		' Comma-Space delimited list of field names
			Dim lsFieldTypes		' Comma-Space delmited list of field types

			' Setup default values (in case procedure failes)
			SetOptions			= False

			mLngFieldTypesAry		= Array()
			mLngPageCount			= -1
			mLngRecordsAffected	= -1

			mStrFieldNamesAry		= Array()
			
			' copy module variables to local ones
			lnAbsolutePage		= mLngAbsolutePage
			lnPageSize			= mLngPageSize
			
			' Reset module variables
			mLngAbsolutePage		= 1
			mLngPageSize			= -1

			If asSQL = "" Then
				Response.Write("<BR><FONT color=red>No SQL provided!!!</FONT><BR>")
				Exit Function
			End If

			' Prepare for errors when executing query
			'On Error Resume Next

			' If the page size has been defined
			If Not lnPageSize	= -1 Then

				' Specify the Microsoft Client cursor
				mObjRs.CursorLocation = ADODB.adUseClient
			
				' Acquire data with ConnectionString
				Call mObjRs.Open(asSQL, mObjConn, ADODB.adOpenKeyset, ADODB.adLockOptimistic)
				
			' Else we are to retrieve all records
			Else

				' Acquire data with ConnectionString
				Call mObjRs.Open(asSQL, mObjConn, ADODB.adOpenStatic, ADODB.adLockBatchOptimistic)
		
			End If ' Not lnPageSize = -1
			
			' Grab the amount of records affected
			mLngRecordsAffected = mObjRs.RecordCount
			
			' If errors occured
			If Err Then
			
				' Append error to internal error collection
				Call DBErr(Err.Description)

				' Clear the error
				Err.Clear
				
				' If the recordset is still open
				If mObjRs.State = ADODB.adStateOpen Then

					' Close the recordset
					Call mObjRs.Close()

				End If ' mObjRs.State = adStateOpen
				
				' Exit the procedure
				Exit Function
			
			' Else if the recordset is not open
			ElseIf Not mObjRs.State = ADODB.adStateOpen Then
				
				' Append error to internal error collection
				Call DBErr("Recordset could not be opened.")
				
				' Exit the procedure
				Exit Function
				
			End If ' Err
			
			' Return positive results
			SetOptions = True
			
			' If user wants to parse meta data
			If ParseMetaData Then
				
				' Loop through each field in the recordset
				For Each loField In mObjRs.Fields
					
					' Append the current field information to Meta-Data strings
					lsFieldNames = lsFieldNames & ", " & loField.Name
					lsFieldTypes = lsFieldTypes & ", " & loField.Type
					
				Next ' loField
				
				' Remove first comma-space delimiter
				lsFieldNames = Mid(lsFieldNames, 3)
				lsFieldTypes = Mid(lsFieldTypes, 3)
				
				' Split Meta Data and store within results
				mStrFieldNamesAry = Split(lsFieldNames, ", ")
				mLngFieldTypesAry = Split(lsFieldTypes, ", ")

			End If ' ParseMetaData
			
			' If data was found
			If Not mObjRs.EOF Then

				' If the page size has not been defined
				If lnPageSize = -1 Then
					
					mLngPageCount = 1

					' Pull data into an array
					asOptions = mObjRs.GetString( , , vbTab, vbCrLf)

				' Else the page size has been defined
				Else
					
					' Define the page size
					mObjRs.PageSize = lnPageSize
					
					' Acquire the number of pages available
					mLngPageCount = mObjRs.PageCount
					
					' Jump to the page that user wants to retrieve
					mObjRs.AbsolutePage = lnAbsolutePage

					' If records still exist
					If Not mObjRs.EOF Then
						
						' Grab records from the current page
						asOptions = mObjRs.GetString( , lnPageSize, vbTab, vbCrLf)

					End If ' Not mObjRs.EOF
					
				End If ' lnPageSize = -1

			End If ' Not mObjRs.EOF

			' If the recordset is open
			If mObjRs.State = ADODB.adStateOpen Then
			
				' Close Recordset
				mObjRs.Close
				
			End If ' mObjRs.State = ADODB.adStateOpen

			If abEncode Then
				asOptions = Server.HTMLEncode(asOptions)
			End If
			
			If Right(asOptions, 2) = vbCrLf Then asOptions = Left(asOptions, Len(asOptions) - 2)
			
			asOptions = Replace(asOptions, vbTab, """>")
			asOptions = Replace(asOptions, vbCrLf, "</OPTION>" & vbCrLf & "<OPTION value=""")

			asOptions = "<OPTION value=""" & asOptions & "</OPTION>" & vbCrLf
			
		End Function ' SetOptions(ByRef asSQL, ByRef asOptions, ByRef abEncode)
' -----------------------------------------------------------------------------
		Public Property Let SQLServer(ByRef abSQLServer)
			
			mBlnIsSQLServer = SQLServer
		
		End Property ' Let SQLServer(ByRef abSQLServer)
' -----------------------------------------------------------------------------
		Public Property Get SQLServer()
			
			SQLServer = mBlnIsSQLServer
		
		End Property ' Get SQLServer()
' -----------------------------------------------------------------------------
		Public Property Get UserID()
			
			UserID = mStrUserID
		
		End Property ' Get UserID()
' -----------------------------------------------------------------------------
		Public Property Let UserID(ByRef asUserID)
		
			mStrUserID = UserID
		
		End Property ' Let UserID(ByRef asUserID)
' -----------------------------------------------------------------------------
		Public Property Get Version()
			
			Version = mStrVersion
		
		End Property ' Get Version()
' -----------------------------------------------------------------------------
	End Class
' -----------------------------------------------------------------------------
'	Set oDatabase = New clsDatabase
'	oDatabase.FieldNames
'	oDatabase.ParseMetaData = True
'	oDatabase.RecordsAffected
'	oDatabase.PageSize
'	oDatabase.PageCount
'	oDatabase.AbsolutePage
%>
