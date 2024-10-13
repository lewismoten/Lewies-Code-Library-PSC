<%
    '**************************************
    ' for :Table Data Export
    '**************************************
    '(c) 2001 Lewis Moten, All rights reserved.
    '**************************************
    ' Name: Table Data Export
    ' Description:Exports data from a databa
    '     se int Transact-SQL Batch files. Used pr
    '     imarily for exporting data from an acces
    '     s database to an SQL Server.
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
    
    ' This script exports data from a databa
    '     se and creates
    ' data to be placed into a Transact-SQL 
    '     Batch File.
    Dim Objconn
    Dim ObjRs
    Dim StrBatch
    Dim StrTableAry
    ' You must modify this line to determine
    '     the tables that you wish to export data 
    '     from
    StrTableAry = Array("Users", "Books", "Products")
    Set ObjConn = createobject("ADODB.Connection")
    Set ObjRs = createobject("ADODB.Recordset")
    ' You must modify this line to reference
    '     the database that the tables are in
    Objconn.Open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=C:\Folder\AccessDatabase.mdb"
    StrBatch = _
    	TableData("Users") & _
    	TableData("Books") & _
    	TableData("Products")
    ObjConn.Close
    Set ObjRs = Nothing
    Set ObjConn = Nothing
    ' Copy & Paste the results from this lin
    '     e into a batch file
    Response.Write "<PRE>" & StrBatch & "</PRE>"
    Function TableData(ByRef pStrTable)
    	
    	Dim lStrSQL
    	Dim lLngMaxIndex
    	Dim lLngIndex
    	Dim lLngVarType
    	
    	lStrSQL = "SELECT * FROM " & pStrTable
    	Set ObjRs = objconn.Execute(lStrSQL)
    	
    	If objRs.EOF Then Exit Function
    	
    	lLngMaxIndex = ObjRs.Fields.Count - 1
    	
    	lStrSQL = "-- Populate " & pStrTable & vbCrLf
    	
    	While Not ObjRs.EOF
    		lStrSQL = lStrSQL & "INSERT INTO [" & pStrTable & "] ("
    		
    		For lLngIndex = 0 To lLngMaxIndex
    			lStrSQL = lStrSQL & "[" & ObjRs.Fields(lLngIndex).Name & "]"
    			If Not lLngIndex = lLngMaxIndex Then lStrSQL = lStrSQL & ", "
    		Next
    		
    		lStrSQL = lStrSQL & ") VALUES ("
    		For lLngIndex = 0 To lLngMaxIndex
    			lVarValue = ObjRs.Fields(lLngIndex).Value
    			lLngVarType = VarType(lVarValue)
    			Select Case lLngVarType
    				Case vbLong, vbInteger, vbDouble, vbSingle, vbDecimal, vbCurrency
    					lStrSQL = lStrSQL & ObjRs.Fields(lLngIndex).Value
    				Case vbNull
    					lStrSQL = lStrSQL & "NULL"
    				Case vbString, vbDate
    					lStrSQL = lStrSQL & "'" & Replace(lVarValue, "'", "''") & "'"
    			End Select
    			If Not lLngIndex = lLngMaxIndex Then lStrSQL = lStrSQL & ", "
    		Next
    		
    		lStrSQL = lStrSQL & ")" & vbCrLf
    		ObjRs.MoveNext
    	Wend
    	TableData = lStrSQL & "GO" & vbCrLf & vbCrLf
    End Function
    %>	
