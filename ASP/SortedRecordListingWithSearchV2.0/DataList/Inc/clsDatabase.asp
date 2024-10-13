<!--METADATA TYPE="TypeLib" NAME="Microsoft ActiveX Data Objects 2.5 Library" UUID="{00000205-0000-0010-8000-00AA006D2EA4}" VERSION="2.5"-->
<SCRIPT language="VBScript" runat="Server" id="LewisMoten.Database">

' Author: Lewis Moten
' Email: Lewis@Moten.com
' URL: http://www.lewismoten.com

' —————————————————————————————————————————————————————————————————————————————
	Class clsDatabase
' —————————————————————————————————————————————————————————————————————————————

		' Define Read/Write properties
		Dim ParseMetaData				' Determines if field names/types are parsed

		' Define internal variables
		Private mbAccessDenied			' Flag determines if access is denied
		Private mbIsSQLServer			' determines if we are talking to an SQL server

		Private mnAbsolutePage			' determines current position of absolute page
		Private mnFieldTypesAry			' Array of Field Types
		Private mnPageCount				' Number of pages returned by recordset
		Private mnPageSize				' determines amount of records are returned on each page
		Private mnVersion				

		Private moConn					' Database connection object
		Private moRs					' Database Recordset

		Private msConnectionString		' Connection string to database
		Private msErrors				' Collection of internal errors
		Private msExpiration			' Date code msExpiration
		Private msFieldNamesAry			' Array of Field Names
		Private msPassword				' Password to access database
		Private mnRecordsAffected		' Amount of records affected by SQL Query
		Private msRegisteredTo			' Who the code is registered to
		Private msUserID				' UserID to login to database with
		Private msVersion				' Version of the code
		
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get ClassName()
			ClassName = "LewisMoten.clsDatabase"
		End Property
' —————————————————————————————————————————————————————————————————————————————
		Private Sub Class_Initialize()
			
			' Initialize exposed properties
			ParseMetaData = False

			' Initialize internal properties
			msFieldNamesAry		= Array()
			mnFieldTypesAry		= Array()
			mnRecordsAffected	= -1
			mnAbsolutePage		= 1
			mnPageSize			= -1
			mnPageCount			= -1
			msExpiration		= ""					' "8/1/2001"
			msRegisteredTo		= "PublicCode"			' "Lewis E. Moten III"
			msVersion			= "2.5"
			
			' Grab database connection information
			msConnectionString	= "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("./") & "\Data\Countries.mdb"		' Application("DBConnectionString")
			msUserID			= ""		' Application("DBUserID")
			msPassword			= ""			' Application("DBPassword")

			mbIsSQLServer		= False					' Application("DBIsSQLServer")
			
			' Make sure class may be used
			mbAccessDenied = False						' AccessDenied()
			
			' Attempt to open database connection
			Call OpenDatabase()
			
		End Sub ' Class_Initialize()
' —————————————————————————————————————————————————————————————————————————————
		Private Sub Class_Terminate()
			
			' Close the database connection
			CloseDatabase()
		
		End Sub ' Class_Terminate()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get AbsolutePage()
			
			AbsolutePage = mnAbsolutePage
		
		End Property ' Get AbsolutePage()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let AbsolutePage(ByRef anAbsolutePage)
			
			mnAbsolutePage = anAbsolutePage
		
		End Property ' Let AbsolutePage()
' —————————————————————————————————————————————————————————————————————————————
		Private Function AccessDenied()
			
			' Assign default return value
			AccessDenied = False
			
			' If the Expiration is a valid date
			If IsDate(msExpiration) Then
				
				' If the Expiration date has all ready occured
				If msExpiration < Date Then
					
					' Append error to internal error collection
					msErrors = msErrors & ";This component has timed out"
					
					' Return positive results
					AccessDenied = True
				
					Call Err.Raise(-1, "AxiomStudios.Database", "Database class has expired.")
					
				End If ' msExpiration < Date
			
			End If ' IsDate(msExpiration)
			
		End Function ' AccessDenied()
' —————————————————————————————————————————————————————————————————————————————
		Public Sub CloseDatabase()
			
			' Prepare for errors
			On Error Resume Next
			
			' If the recordset object is defined
			If IsObject(moRs) Then
					
				' If the recordset object is open
				If moRs.State = ADODB.adStateOpen Then
						
					' Close Recordset
					moRs.Close
					
				End If ' moRs.State = ADODB.adStateOpen
					
				' Release object from memory
				Set moRs = Nothing

			End If ' IsObject(moRs)
			
			' If connection object has been created
			If IsObject(moConn) Then
				
				' If the connection object is open
				If moConn.State = ADODB.adStateOpen Then
					
					' Close the connection
					moConn.Close
				
				End If ' moConn.State = ADODB.adStateOpen
				
				' Release object from memory
				Set moConn = Nothing
				
			End If ' IsObject(moConn)
			
		End Sub ' CloseDatabase()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get ConnectionString()

			ConnectionString = msConnectionString
		
		End Property ' Get ConnectionString()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let ConnectionString(ByRef asConnectionString)
		
			msConnectionString = ConnectionString
		
		End Property ' Let ConnectionString(ByRef asConnectionString)
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get Errors()

'			Dim lnIndex			' Index of connection errors
'			Dim lsConnErrors	' Collection of connection errors
'						
'			' Initialize index as not defined
'			lnIndex = -1
'
'			' If the connection object was defined			
'			If IsObject(moConn) Then
'				
'				' If error exist within the conneciton object
'				If Not moConn.Errors.Count = 0 Then			
'						
'					' Loop through each error found within the connection object
'					For lnIndex = 0 To moConn.Errors.Count -1
'							
'						If Not lnIndex = 0 Then lsConnErrors = lsConnmsErrors & ";"
'							
'						' Append the connection error to the log
'						lsConnErrors = lsConnErrors & _
'							"ADODB.Connection.Errors: " & _
'							moConn.Errors.Item(lnIndex).Description
'
'					Next ' lnIndex
'						
'				End If ' Not moConn.Errors.Count = 0
'
'			End If ' IsObject(moConn)
			

			' If no errors are defined
			If Len(msErrors) <= 1 Then
			
				' Do nothing
			
			' Else errors are present
			Else
			
				' If errors begin with a delimiter
				If Left(msErrors, 1) = ";" Then
					
					' Return the errors with the first character removed
					Errors = Mid(msErrors, 2)
			
				' Else errors do not begin with a delimiter
				Else
					
					' Return errors
					Errors = msErrors
				
				End If ' Left(msErrors, 1) = ";"
				
			End If ' Len(msErrors) <= 1
			
'			If Errors = "" Then
'				If lsConnErrors = "" Then
'					Errors = ""
'				Else
'					Errors = lsConnErrors
'				End If
'			Else
'				If lsConnErrors = "" Then
'					' do nothing
'				Else
'					Errors = Errors & ";" & lsConnErrors
'				End If
'			End If
				
			
		End Property ' Get Errors()
' —————————————————————————————————————————————————————————————————————————————
		Public Function ExecuteSQL(ByRef asSQL)
		
			' Setup default values (in case procedure failes)
			mnRecordsAffected = -1
			ExecuteSQL = False

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

			' Prepare for errors when executing query
			On Error Resume Next

			Call moConn.Execute(asSQL, mnRecordsAffected, ADODB.adCmdText Or ADODB.adExecuteNoRecords)
		
			' If errors occured
			If Err Then

				' Append error to internal error collection
				msErrors = msErrors & ";" & Err.Description

				' Clear the error
				Err.Clear
			
			' Else SQL query executed fine
			Else
				
				' Return positive results
				ExecuteSQL = True
				
			End If ' Err
		
		End Function ' ExecuteSQL(ByRef asSQL)
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get Expiration()
			
			' If the Expiration date is a valid date
			If IsDate(msExpiration) Then

				' Return the date that this class msExpiration
				Expiration = msExpiration

			' Else the class never msExpiration
			Else

				' Let the user know that the class never msExpiration
				Expiration = "Never"

			End If ' IsDate(msExpiration)
			
		End Property ' Expiration()
' —————————————————————————————————————————————————————————————————————————————
		Public Function FieldNames()
		
			FieldNames = msFieldNamesAry
			
		End Function ' FieldNames()
' —————————————————————————————————————————————————————————————————————————————
		Public Function FieldTypes()
			
			Dim lnIndex		' Index of Field Types Array
			Dim lnValue		' Value of number within field types array
			
			' Loop through each index in the field types array
			For lnIndex = 0 To UBound(mnFieldTypesAry, 1)
			
				' Grab the value of the current index
				lnValue = mnFieldTypesAry(lnIndex)
				
				' If the value is a number
				If IsNumeric(lnValue) And Not lnValue = "" Then
					
					' Convert SubType String to a SubType Long
					mnFieldTypesAry(lnIndex) = CLng(lnValue)
				
				End If ' IsNumeric(lnValue) And Not lnValue = ""
				
			Next ' lnIndex
			
			' Return the values of the field types
			FieldTypes = mnFieldTypesAry
			
		End Function ' FieldTypes()
' —————————————————————————————————————————————————————————————————————————————
		Public Function FormatBoolean(ByVal abExpression, ByRef abDefault)

			Dim lbTrue
			Dim lbFalse
			
			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

			' If database is an SQL Server
			If mbIsSQLServer Then
			
				' Assign value for true
				lbTrue = 1
			
			' Else database is an Access database
			Else
				
				' Assign value for true
				lbTrue = -1
			
			End If ' mbIsSQLServer
			
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
' —————————————————————————————————————————————————————————————————————————————
		Public Function FormatDate(ByRef adExpression)

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

			' If Expression valid date
			If IsDate(adExpression) Then

				' If database is an SQL Server
				If mbIsSQLServer Then
				
					' Return formated date for SQL Server
					FormatDate = "'" & adExpression & "'"
				
				' Else database is an Access File
				Else
					
					' Return formated date for Access File
					FormatDate = "#" & adExpression & "#"
				
				End If ' mbIsSQLServer

			' Else Expression not valid date
			Else

				' Return NULL
				FormatDate = "NULL"

			End If ' IsDate(adExpression)

		End Function ' FormatDate(ByRef adExpression)
' —————————————————————————————————————————————————————————————————————————————
		Public Function FormatMoney(ByVal acExpression)

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

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
' —————————————————————————————————————————————————————————————————————————————
		Public Function FormatNumeric(ByVal anExpression)

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

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
' —————————————————————————————————————————————————————————————————————————————
		Public Function FormatString(ByVal asExpression, ByRef anMaxLength)

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function

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
' —————————————————————————————————————————————————————————————————————————————
		Public Function OpenDatabase()

			' Setup default value as false
			OpenDatabase = False
			
			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function
			
			' If a connection string has not been defined
			If msConnectionString = "" Then
				
				' Append error to error collection
				msErrors = msErrors & ";No connection string has been defined."
			
				' Exit the procedure
				Exit Function
				
			End If ' msConnectionString = ""

			' If a previouse connection object was created
			If IsObject(moConn) Then
				
				' If an open database connection exists
				If moConn.State = ADODB.adStateOpen Then
					
					' Close current database connection
					moConn.Close
				
				End If ' moConn.State = ADODB.adStateOpen
			
			' Else no connection object has been created yet
			Else
			
				' Create Connection Object
				Set moConn = Server.CreateObject("ADODB.Connection")
				
			End If ' IsObject(moConn)

			' Prepare for errors
			On Error Resume Next
			
			' Open database connection
			Call moConn.Open(msConnectionString, msUserID, msPassword)

			' If any errors have occured
			If Err Then
				
				If Err Then
					' Append error to internal error collection
					msErrors = msErrors & ";" & Err.Description
					
					' Clear errors
					Err.Clear
				End If
				
				' Release connection object
				Set moConn = Nothing

			' Else no errrors occured.
			Else
			
				' If the recordset has not been opened
				If Not IsObject(moRs) Then
				
					' Create a recordset object
					Set moRs = Server.CreateObject("ADODB.Recordset")
				
				End If ' Not IsObject(moRs)
				
				' Return Positive Results
				OpenDatabase = True
				
			End If ' Err

		End Function ' OpenDatabase()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get PageCount()
			
			PageCount = mnPageCount
		
		End Property ' Get PageCount()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get PageSize()
			
			PageSize = mnPageSize
		
		End Property ' Get PageSize()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let PageSize(ByRef anPageSize)
			
			mnPageSize = anPageSize
		
		End Property
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get Password()
			
			Password = msPassword
		
		End Property ' Get Password()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let Password(ByRef asPassword)
			
			msPassword = Password
		
		End Property ' Let Password(ByRef asPassword)
' —————————————————————————————————————————————————————————————————————————————
		Public Function QueryFields(ByRef asFieldAry, ByVal asQuery)

			' Bug discovered by Lucas Moten 10/16/2000
			' query containing =^% or other non alpha-numeric characters
			' is causing the query not to pass back anything
			
			Dim loRegExp			' Regular Expression Object (requires vbScript 5.0)
			Dim loRequiredWords		' Words that MUST match within a search
			Dim loUnwantedWords		' Words that MUST NOT match within a search
			Dim loOptionalWords		' Words that AT LEAST ONE must match
			Dim lsSQL				' Arguments of SQL query that is returned (WHERE __Arguments___)
			Dim lnIndex				' Index of an array
			Dim lsPhrase			' Keyword or Phrase being worked with
			Dim lsExpressionPhrase
			Dim lsExpressionWord
			
			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function
			
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
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get RecordsAffected()
			
			' Return the amount of records affected by last SQL query
			RecordsAffected = mnRecordsAffected
			
		End Property ' Get RecordsAffected()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get RegisteredTo()
		
			' Return the value of the name that the class is registered to
			RegisteredTo = msRegisteredTo
			
		End Property ' Get RegisteredTo()
' —————————————————————————————————————————————————————————————————————————————
		Public Function SetData(ByRef asSQL, ByRef avDataAry)

			Dim lnAbsolutePage		' Page number of record sets
			Dim lnPageSize			' Number of records on each page
			Dim lnPage				' Current page being worked with

			Dim lsFieldNames		' Comma-Space delimited list of field names
			Dim lsFieldTypes		' Comma-Space delmited list of field types

			' Setup default values (in case procedure failes)
			SetData				= False

			mnFieldTypesAry		= Array()
			mnPageCount			= -1
			mnRecordsAffected	= -1

			msFieldNamesAry		= Array()
			
			' copy module variables to local ones
			lnAbsolutePage		= mnAbsolutePage
			lnPageSize			= mnPageSize
			
			' Reset module variables
			mnAbsolutePage		= 1
			mnPageSize			= -1

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function
			
			If asSQL = "" Then
				Response.Write("<BR><FONT color=red>No SQL provided!!!</FONT><BR>")
				Exit Function
			End If

			' Prepare for errors when executing query
			'On Error Resume Next

'			if moconn.Errors then
'			for i = 0 to moconn.Errors.Count - 1
'			Response.Write "<LI>" & moconn.Errors(i) & "</LI>"
'			next
'			end if

			' If the page size has been defined
			If Not lnPageSize	= -1 Then

				' Specify the Microsoft Client cursor
				moRs.CursorLocation = ADODB.adUseClient
			
				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, moConn, ADODB.adOpenKeyset, ADODB.adLockOptimistic)
				
			' Else we are to retrieve all records
			Else

				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, moConn, ADODB.adOpenStatic, ADODB.adLockBatchOptimistic)
		
			End If ' Not lnPageSize = -1
			
			' Grab the amount of records affected
			mnRecordsAffected = moRs.RecordCount
			
			' If errors occured
			If Err Then
			
				' Append error to internal error collection
				msErrors = msErrors & ";" & Err.Description

				' Clear the error
				Err.Clear
				
				' If the recordset is still open
				If moRs.State = ADODB.adStateOpen Then

					' Close the recordset
					Call moRs.Close()

				End If ' moRs.State = adStateOpen
				
				' Exit the procedure
				Exit Function
			
			' Else if the recordset is not open
			ElseIf Not moRs.State = ADODB.adStateOpen Then
				
				' Append error to internal error collection
				msErrors = msErrors & ";Recordset could not be opened."
				
				' Exit the procedure
				Exit Function
				
			End If ' Err
			
			' Return positive results
			SetData = True
			
			' If user wants to parse meta data
			If ParseMetaData Then
				
				' Loop through each field in the recordset
				For Each loField In moRs.Fields
					
					' Append the current field information to Meta-Data strings
					lsFieldNames = lsFieldNames & ", " & loField.Name
					lsFieldTypes = lsFieldTypes & ", " & loField.Type
					
				Next ' loField
				
				' Remove first comma-space delimiter
				lsFieldNames = Mid(lsFieldNames, 3)
				lsFieldTypes = Mid(lsFieldTypes, 3)
				
				' Split Meta Data and store within results
				msFieldNamesAry = Split(lsFieldNames, ", ")
				mnFieldTypesAry = Split(lsFieldTypes, ", ")

			End If ' ParseMetaData
			
			' If data was found
			If Not moRs.EOF Then

				' If the page size has not been defined
				If lnPageSize = -1 Then
					
					mnPageCount = 1

'##TEST##					
'					dim field
'					Response.Write("<table border=1><TR>")
'					for each field in mors.Fields
'						Response.Write "<TD>" & field.name & "</TD>"
'					next
'					Response.Write("</TR>")
'					while not mors.EOF
'					Response.Write("<TR>")
'					for each field in mors.Fields
'						Response.Write "<TD>" & field.value & "</TD>"
'					next
'					Response.Write("</TR>")
'					mors.MoveNext
'					wend
'					Response.Write("</TABLE>")
'
					' Pull data into an array
					avDataAry = moRS.GetRows

				' Else the page size has been defined
				Else
					
					' Define the page size
					moRs.PageSize = lnPageSize
					
					' Acquire the number of pages available
					mnPageCount = moRs.PageCount
					
					' Jump to the page that user wants to retrieve
					moRs.AbsolutePage = lnAbsolutePage

					' If records still exist
					If Not moRs.EOF Then
						
						' Grab records from the current page
						avDataAry = moRs.GetRows(lnPageSize, ADODB.adBookmarkCurrent)

					End If ' Not moRs.EOF
					
				End If ' lnPageSize = -1

			End If ' Not moRS.EOF

			' If the recordset is open
			If moRs.State = ADODB.adStateOpen Then
			
				' Close Recordset
				moRS.Close
				
			End If ' moRs.State = ADODB.adStateOpen

		End Function ' SetData(ByRef asSQL, ByRef avDataAry)
' —————————————————————————————————————————————————————————————————————————————
		Public Function SetOptions(ByRef asSQL, ByRef asOptions, ByRef abEncode)

			Dim lnAbsolutePage		' Page number of record sets
			Dim lnPageSize			' Number of records on each page
			Dim lnPage				' Current page being worked with

			Dim lsFieldNames		' Comma-Space delimited list of field names
			Dim lsFieldTypes		' Comma-Space delmited list of field types

			' Setup default values (in case procedure failes)
			SetOptions			= False

			mnFieldTypesAry		= Array()
			mnPageCount			= -1
			mnRecordsAffected	= -1

			msFieldNamesAry		= Array()
			
			' copy module variables to local ones
			lnAbsolutePage		= mnAbsolutePage
			lnPageSize			= mnPageSize
			
			' Reset module variables
			mnAbsolutePage		= 1
			mnPageSize			= -1

			' Cancel process if access is denied
			If mbAccessDenied Then Exit Function
			
			If asSQL = "" Then
				Response.Write("<BR><FONT color=red>No SQL provided!!!</FONT><BR>")
				Exit Function
			End If

			' Prepare for errors when executing query
			'On Error Resume Next

'			if moconn.Errors then
'			for i = 0 to moconn.Errors.Count - 1
'			Response.Write "<LI>" & moconn.Errors(i) & "</LI>"
'			next
'			end if

			' If the page size has been defined
			If Not lnPageSize	= -1 Then

				' Specify the Microsoft Client cursor
				moRs.CursorLocation = ADODB.adUseClient
			
				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, moConn, ADODB.adOpenKeyset, ADODB.adLockOptimistic)
				
			' Else we are to retrieve all records
			Else

				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, moConn, ADODB.adOpenStatic, ADODB.adLockBatchOptimistic)
		
			End If ' Not lnPageSize = -1
			
			' Grab the amount of records affected
			mnRecordsAffected = moRs.RecordCount
			
			' If errors occured
			If Err Then
			
				' Append error to internal error collection
				msErrors = msErrors & ";" & Err.Description

				' Clear the error
				Err.Clear
				
				' If the recordset is still open
				If moRs.State = ADODB.adStateOpen Then

					' Close the recordset
					Call moRs.Close()

				End If ' moRs.State = adStateOpen
				
				' Exit the procedure
				Exit Function
			
			' Else if the recordset is not open
			ElseIf Not moRs.State = ADODB.adStateOpen Then
				
				' Append error to internal error collection
				msErrors = msErrors & ";Recordset could not be opened."
				
				' Exit the procedure
				Exit Function
				
			End If ' Err
			
			' Return positive results
			SetOptions = True
			
			' If user wants to parse meta data
			If ParseMetaData Then
				
				' Loop through each field in the recordset
				For Each loField In moRs.Fields
					
					' Append the current field information to Meta-Data strings
					lsFieldNames = lsFieldNames & ", " & loField.Name
					lsFieldTypes = lsFieldTypes & ", " & loField.Type
					
				Next ' loField
				
				' Remove first comma-space delimiter
				lsFieldNames = Mid(lsFieldNames, 3)
				lsFieldTypes = Mid(lsFieldTypes, 3)
				
				' Split Meta Data and store within results
				msFieldNamesAry = Split(lsFieldNames, ", ")
				mnFieldTypesAry = Split(lsFieldTypes, ", ")

			End If ' ParseMetaData
			
			' If data was found
			If Not moRs.EOF Then

				' If the page size has not been defined
				If lnPageSize = -1 Then
					
					mnPageCount = 1

					' Pull data into an array
					asOptions = moRs.GetString( , , vbTab, vbCrLf)

				' Else the page size has been defined
				Else
					
					' Define the page size
					moRs.PageSize = lnPageSize
					
					' Acquire the number of pages available
					mnPageCount = moRs.PageCount
					
					' Jump to the page that user wants to retrieve
					moRs.AbsolutePage = lnAbsolutePage

					' If records still exist
					If Not moRs.EOF Then
						
						' Grab records from the current page
						asOptions = moRs.GetString( , lnPageSize, vbTab, vbCrLf)

					End If ' Not moRs.EOF
					
				End If ' lnPageSize = -1

			End If ' Not moRS.EOF

			' If the recordset is open
			If moRs.State = ADODB.adStateOpen Then
			
				' Close Recordset
				moRS.Close
				
			End If ' moRs.State = ADODB.adStateOpen

			If abEncode Then
				asOptions = Server.HTMLEncode(asOptions)
			End If
			
			If Right(asOptions, 2) = vbCrLf Then asOptions = Left(asOptions, Len(asOptions) - 2)
			
			asOptions = Replace(asOptions, vbTab, """>")
			asOptions = Replace(asOptions, vbCrLf, "</OPTION>" & vbCrLf & "<OPTION value=""")

			asOptions = "<OPTION value=""" & asOptions & "</OPTION>" & vbCrLf
			
		End Function ' SetOptions(ByRef asSQL, ByRef asOptions, ByRef abEncode)
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let SQLServer(ByRef abSQLServer)
			
			mbIsSQLServer = SQLServer
		
		End Property ' Let SQLServer(ByRef abSQLServer)
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get SQLServer()
			
			SQLServer = mbIsSQLServer
		
		End Property ' Get SQLServer()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get UserID()
			
			UserID = msUserID
		
		End Property ' Get UserID()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let UserID(ByRef asUserID)
		
			msUserID = UserID
		
		End Property ' Let UserID(ByRef asUserID)
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get Version()
			
			Version = msVersion
		
		End Property ' Get Version()
' —————————————————————————————————————————————————————————————————————————————
	End Class
' —————————————————————————————————————————————————————————————————————————————
'	Set oDatabase = New clsDatabase
'	oDatabase.PageSize
'	oDatabase.PageCount
'	oDatabase.AbsolutePage
</SCRIPT>
