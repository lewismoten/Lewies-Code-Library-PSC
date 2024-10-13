<%
    '**************************************
    ' for :Common Database Routines
    '**************************************
    'Copyright (c) 1999 by Lewis Moten, All rights reserved.
    '**************************************
    ' Name: Common Database Routines
    ' Description:To assist in interfacing w
    '     ith databases. This script can format va
    '     riables and return SQL formats. Such as 
    '     double quoting apposterphies and surroun
    '     ding strings with quotes, Returning NULL
    '     for invalid data types, trimming strings
    '     so they do not exceed maximum lengths. T
    '     his also has some functions so that you 
    '     can open and close databases more convei
    '     ently with just one line of code. You ca
    '     n query a database and get an Array as w
    '     ell with some code.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:This script assumes that you at
    '     least have Microsoft ActiveX Data Object
    '     s 2.0 or Higher (ADODB). This script may
    '     get some getting used to at first until 
    '     you go through and study what each routi
    '     ne can do.
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
%>
    
    <!--METADATA Type="TypeLib" NAME="Microsoft ActiveX Data Objects 2.0 Library" UUID="{00000200-0000-0010-8000-00AA006D2EA4}" VERSION="2.0"-->
    <%
    ' Setup the ConnectionString
    Dim sCONNECTION_STRING
    sCONNECTION_STRING = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=D:\inetpub\wwwroot\inc\data\database.mdb;"
    Dim oConn
    '---------------------------------------
    '     ----------------------------------------
    '     
    Function DBConnOpen(ByRef aoConnObj)
    	' This routine connects to a database and returns
    	' weather or not it was successful
    	' Prepare for any errors that may occur while connecting to the database
    	On Error Resume Next
    	' Create a connection object
    	Set aoConnObj = Server.CreateObject("ADODB.Connection")
    	' Open a connection to the database
    	Call aoConnObj.Open(sCONNECTION_STRING)
    	' If any errors have occured
    	If Err Then
    		' Clear errors
    		Err.Clear
    		' Release connection object
    		Set aoConnObj = Nothing
    		' Return unsuccessful results
    		DBConnOpen = False
    	' Else errors did not occur
    	Else
    		' Return successful results
    		DBConnOpen = True
    	End If ' Err
    End Function ' DBConnOpen
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function DBConnClose(ByRef aoConnObj)
    	' This routine closes the database connection and releases objects
    	' from memory
    	' If the connection variable has been defined as an object
    	If IsObject(aoConnObj) Then
    		' If the connection is open
    		If aoConnObj.State = adStateOpen Then
    			' Close the connection
    			aoConnObj.Close
    			' Return positive Results
    			DBConnClose = True
    		End If ' aoConnObj.State = adStateOpen
    		' Release connection object
    		Set aoConnObj = Nothing
    	End If ' IsObject(aoConnObj)
    End Function ' DBConnClose
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SetData(ByRef asSQL, ByRef avDataAry)
    	' This routine acquires data from the database
    	Dim loRS ' ADODB.Recordset Object
    	' Create Recordset Object
    	Set loRS = Server.CreateObject("ADODB.Recordset")
    	' Prepare for errors when opening database connection
    	On Error Resume Next
    	' If a connection object has been defined
    	If IsObject(oConn) Then
    		' If the connection is open
    		If oConn.State = adStateOpen Then
    			' Acquire data with connection object
    			Call loRS.Open(asSQL, oConn, adOpenForwardOnly, adLockReadOnly)
    		' Else the connection is closed
    		Else
    			' Set the ConnectionString
    			Call SetConnectionString(csConnectionString)
    			' If atempt to open connection succeeded
    			If DBConnOpen() Then
    				' Acquire data with connection object
    				Call loRS.Open(asSQL, oConn, adOpenForwardOnly, adLockReadOnly)
    				' Return connection object to closed state
    				Call DBConnClose()
    			End If ' DBConnOpen()
    		End If ' aoConn.State = adStateOpen
    	' Else active connection is the ConnectionString
    	Else
    		' Acquire data with ConnectionString
    		Call loRS.Open(asSQL, sCONNECTION_STRING, adOpenForwardOnly, adLockReadOnly)
    	End If ' IsObject(oConn)
    	' If errors occured
    	If Err Then
    		response.write "<HR color=red>" & err.description & "<HR color=red>" & asSQL & "<HR color=red>"
    		' Clear the error
    		Err.Clear
    		' If the recorset is open
    		If loRS.State = adStateOpen Then
    			' Close the recorset
    			loRS.Close
    		End If ' loRS.State = adStateOpen
    		' Release Recordset from memory
    		Set loRS = Nothing
    		' Return negative results
    		SetData = False
    		' Exit Routine
    		Exit Function
    	End If ' Err
    	' Return positve results
    	SetData = True
    	' If data was found
    	If Not loRS.EOF Then
    		' Pull data into an array
    		avDataAry = loRS.GetRows
    	End If ' Not loRS.EOF
    	' Close Recordset
    	loRS.Close
    	' Release object from memory
    	Set loRS = Nothing
    End Function ' SetData
    '---------------------------------------
    '     ----------------------------------------
    '     
    ' SQL Preperations are used to prepare v
    '     ariables for SQL Queries. If
    ' invalid data is passed to these routin
    '     es, NULL values or Default Data
    ' is returned to keep your SQL Queries f
    '     rom breaking from users breaking
    ' datatype rules.
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SQLPrep_s(ByVal asExpression, ByRef anMaxLength)
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
    		SQLPrep_s = "NULL"
    	' Else expression is not empty
    	Else
    		' Return quoted expression
    		SQLPrep_s = "'" & asExpression & "'"
    	End If ' asExpression
    End Function ' SQLPrep_s
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SQLPrep_n(ByVal anExpression)
    	' If expression numeric
    	If IsNumeric(anExpression) And Not anExpression = "" Then
    		' Return number
    		SQLPrep_n = anExpression
    	' Else expression not numeric
    	Else
    		' Return NULL
    		SQLPrep_n = "NULL"
    	End If ' IsNumeric(anExpression) And Not anExpression = ""
    End Function ' SQLPrep_n
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SQLPrep_b(ByVal abExpression, ByRef abDefault)
    	' Declare Database Constants
    	Const lbTRUE = -1 '1 = SQL, -1 = Access
    	Const lbFALSE = 0
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
    		SQLPrep_b = lbTRUE
    	' Else Result is false
    	Else
    		' Return False
    		SQLPrep_b = lbFALSE
    	End If ' lbResult
    End Function ' SQLPrep_b
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SQLPrep_d(ByRef adExpression)
    	' If Expression valid date
    	If IsDate(adExpression) Then
    		' Return Date
    		'SQLPrep_d = "'" & adExpression & "'" ' SQL Database
    		SQLPrep_d = "#" & adExpression & "#" ' Access Database
    	' Else Expression not valid date
    	Else
    		' Return NULL
    		SQLPrep_d = "NULL"
    	End If ' IsDate(adExpression)
    End Function ' SQLPrep_d
    '---------------------------------------
    '     ----------------------------------------
    '     
    Public Function SQLPrep_c(ByVal acExpression)
    	' If Empty Expression
    	If acExpression = "" Then
    		' Return Null
    		SQLPrep_c = "NULL"
    	' Else expression has content
    	Else
    		' Prepare for Errors
    		On Error Resume Next
    		' Attempt to convert expression to currency
    		SQLPRep_c = CCur(acExpression)
    		' If error occured
    		If Err Then
    			' Clear Error
    			Err.Clear
    			SQLPrep_c = "NULL"
    		End If ' Err
    	End If ' acExpression = ""
    End Function ' SQLPrep_c
    '---------------------------------------
    '     ----------------------------------------
    '     
    Function buildJoinStatment(sTable,sFldLstAry,rs,conn)
    Dim i,sSql,sTablesAry,sJnFldsAry,bJoinAry,sJoinDisplay
    ReDim sTablesAry(UBound(sFldLstAry))
    ReDim sJnFldsAry(UBound(sFldLstAry))
    ReDim bJoinAry(UBound(sFldLstAry))
    For i = 0 To UBound(sFldLstAry)
    sSql = "SELECT OBJECT_NAME(rkeyid),COL_NAME(rkeyid,rkey1)"
    sSql = sSql &" FROM sysreferences"
    sSql = sSql &" WHERE fkeyid = OBJECT_ID('"& sTable &"') "
    sSql = sSql &" AND col_name(fkeyid,fkey1) = '"& Trim(sFldLstAry(i)) &"'"
    rs.open sSql,conn
    If Not rs.eof Then
    sTablesAry(i) = rs(0)
    sJnFldsAry(i) = rs(1)
    End If
    rs.close
    Next
    If UBound(sFldLstAry) >= 0 Then
    For i = 0 To UBound(sFldLstAry)
    If sTablesAry(i) <> "" Then
    bJoinAry(i) = True
    Else
    bJoinAry(i) = False
    End If
    If i <> UBound(sFldLstAry) Then sSql = sSql &" +' - '+ "
    Next
    sSql = "FROM "& sTable
    For i = 0 To UBound(sFldLstAry)
    If bJoinAry(i) Then sSql = sSql &" LEFT JOIN "& sTablesAry(i) &" ON "& sTable &"."& sFldLstAry(i) &" = "& sTablesAry(i) &"."& sJnFldsAry(i)
    Next
    End If
    buildJoinStatment = sSql
    End Function
    '---------------------------------------
    '     ----------------------------------------
    '     
    Function buildQuery(ByRef asFieldAry, ByVal asKeyWords)
    	' To find fields that may have a word in them
    	' OR roger
    	' | roger
    	' roger
    	' To find fields that must match a word
    	' AND roger
    	' + roger
    	' & roger
    	' To find fields that must not match a word
    	' NOT roger
    	' - roger
    	' Also use phrases
    	' +"rogers dog" -cat
    	' +(rogers dog)
    	Dim loRegExp
    	Dim loRequiredWords
    	Dim loUnwantedWords
    	Dim loOptionalWords
    	Dim lsSQL
    	Dim lnIndex
    	Dim lsKeyword
    	Set loRegExp = New RegExp
    	loRegExp.Global = True
    	loRegExp.IgnoreCase = True
    	loRegExp.Pattern = "((AND|[+&])\s*[\(\[\{""].*[\)\]\}""])|((AND\s|[+&])\s*\b[-\w']+\b)"
    	Set loRequiredWords = loRegExp.Execute(asKeywords)
    	asKeywords = loRegExp.Replace(asKeywords, "")
    	loRegExp.Pattern = "(((NOT|[-])\s*)?[\(\[\{""].*[\)\]\}""])|(((NOT\s+|[-])\s*)\b[-\w']+\b)"
    	Set loUnwantedWords = loRegExp.Execute(asKeywords)
    	asKeywords = loRegExp.Replace(asKeywords, "")
    	loRegExp.Pattern = "(((OR|[|])\s*)?[\(\[\{""].*[\)\]\}""])|(((OR\s+|[|])\s*)?\b[-\w']+\b)"
    	Set loOptionalWords = loRegExp.Execute(asKeywords)
    	asKeywords = loRegExp.Replace(asKeywords, "")
    	If Not loRequiredWords.Count = 0 Then
    		' REQUIRED
    		lsSQL = lsSQL & "("
    		For lnIndex = 0 To loRequiredWords.Count - 1
    			lsKeyword = loRequiredWords.Item(lnIndex).Value
    			loRegExp.Pattern = "^(AND|[+&])\s*"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			loRegExp.Pattern = "[()""\[\]{}]"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			lsKeyword = Replace(lsKeyword, "'", "''")
    			If Not lnIndex = 0 Then
    				lsSQL = lsSQL & " AND "
    		 	End If
    			lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsKeyword & "%' OR ") & " LIKE '%" & lsKeyword & "%')"
    		Next
    		lsSQL = lsSQL & ")"
    	End If
    	If Not loOptionalWords.Count = 0 Then
    		' OPTIONAL
    		If lsSQL = "" Then
    			lsSQL = lsSQL & "("
    		Else
    			lsSQL = lsSQL & " AND ("
    		End If
    		For lnIndex = 0 To loOptionalWords.Count - 1
    			lsKeyword = loOptionalWords.Item(lnIndex).Value
    			loRegExp.Pattern = "^(OR|[|])\s*"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			loRegExp.Pattern = "[()""\[\]{}]"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			lsKeyword = Replace(lsKeyword, "'", "''")
    			If Not lnIndex = 0 Then
    				lsSQL = lsSQL & " OR "
    			End If
    			lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsKeyword & "%' OR ") & " LIKE '%" & lsKeyword & "%')"
    		Next
    		lsSQL = lsSQL & ")"
    	End If
    	If Not loUnwantedWords.Count = 0 Then
    		' UNWANTED
    		If lsSQL = "" Then
    			lsSQL = lsSQL & "NOT ("
    		Else
    			lsSQL = lsSQL & " AND NOT ("
    		End If
    		For lnIndex = 0 To loUnwantedWords.Count - 1
    			lsKeyword = loUnWantedWords.Item(lnIndex).Value
    			loRegExp.Pattern = "^(NOT|[-])\s*"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			loRegExp.Pattern = "[()""\[\]{}]"
    			lsKeyword = loRegExp.Replace(lsKeyword, "")
    			lsKeyword = Replace(lsKeyword, "'", "''")
    			If Not lnIndex = 0 Then
    				lsSQL = lsSQL & " OR "
    			End If
    			lsSQL = lsSQL & "(" & Join(asFieldAry, " LIKE '%" & lsKeyword & "%' OR ") & " LIKE '%" & lsKeyword & "%')"
    		Next
    		lsSQL = lsSQL & ")"
    	End If
    	If Not lsSQL = "" Then lsSQL = "(" & lsSQL & ")"
    	buildQuery = lsSQL
    End Function
    '---------------------------------------
    '     ----------------------------------------
    '     
    %>