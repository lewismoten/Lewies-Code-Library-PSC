<%
' NOTICE:
'	This file is dependent upon clsDatabase.asp
'	Please make sure you have referenced this file on the page you
'	are creating the list on.

' Author: Lewis Moten
' Email: Lewis@Moten.com
' URL: http://www.lewismoten.com

' ------------------------------------------------------------------------------
Class clsDataList
' ------------------------------------------------------------------------------
	
	Public AllowSorting					' Allow users to sort fields?
	Public TableName					' name of table to request data from
	Public PageSize						' Number of records on each page
	Public Arguments					' Additional arguments for where clause
										' WHERE ___Arguments___ and 1 = 2 and ...
	Public ImageDirectory				' Where the images are located
	Public DefaultSortedField			' default field data is sorted against
	Public DefaultSortedOrder			' default order data is sorted against
	Public ListOrderField				' Field signifying list order
	Public PrimaryKey					' Field identifying record

	Private msSearchFields
	Private msReturnFields
	Private msReturnFieldCaptions	
	
	Private msSearchFieldAry			' Fields to search
	Private msReturnFieldAry			' Fields to be returned
	Private msReturnFieldCaptionAry		' Caption to display for returned fields
	
	Private mnReturnFieldCount			' Number of fields to be returned
	Private mnSearchFieldCount			' Number of fields to be searched
	
	Private msURLPrefix					' URL prefix is used to create other URL's
	Private msSortedField				' Current field data is sorted against
	Private msSortedOrder				' Current order that data is sorted against
	Private msQuery						' Query user is searching for
	Private mnPage						' Page being viewed
	Private mnLastPage					' Last page of results
	Private mnRecordCount				' Number of records found
	
	Public urlAdd
	Public urlEdit
	Public urlDelete
	Public urlMoveUp
	Public urlMoveDown
	Public urlInsert
	'Public urlList
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		mnReturnFieldCount = 0
		mnSearchFieldCount = 0
		msSortedField = Request.QueryString("SortedField")
		msSortedOrder = Request.QueryString("SortedOrder")
		msQuery = Request.QueryString("Query")
		mnPage = Request.QueryString("Page")
		If mnPage = "" Or Not IsNumeric(mnPage) Then mnPage = 1 Else mnPage = CLng(mnPage)
		If mnPage < 1 Then mnPage = 1
		PageSize = 10
		AllowSorting = True
		ImageDirectory = "images/"
	End Sub
' ------------------------------------------------------------------------------
	Private Function SortQuery(ByRef psFieldName)

		Dim lsQueryString
		
		lsQueryString = Request.QueryString

		' Remove Sorting properties
		lsQueryString = Replace(lsQueryString, "SortedField=" & Server.URLEncode(msSortedField), "")
		lsQueryString = Replace(lsQueryString, "SortedOrder=" & Server.URLEncode(msSortedOrder), "")
		
		' Remove double ampersants
		While Not InStr(1, lsQueryString, "&&") = 0
			lsQueryString = Replace(lsQueryString, "&&", "&")
		Wend
		
		' Remove ampersant if it is the only character.
		If Len(lsQueryString) = 1 Then
			lsQueryString = ""
		Else
			If Not Right(lsQueryString, 1) = "&" Then
				lsQueryString = lsQueryString & "&"
			End If
			
			' Remove ampersant if it is the first character
			If Left(lsQueryString, 1) = lsQueryString Then
				lsQueryString = Mid(lsQueryString, 2)
			End If
		End If
		
		lsQueryString = lsQueryString & "SortedField=" & _
			Server.URLEncode(psFieldName) & _
			"&SortedOrder="

		If msSortedField = psFieldName Then
			Select Case msSortedOrder
				Case "ASC"
					lsQueryString = lsQueryString & "DESC"
				Case "DESC"
					'do nothing
				Case Else
					lsQueryString = lsQueryString & "ASC"
			End Select
		ElseIf msSortedField = "" Then
			If psFieldName = DefaultSortedField Then
				Select Case DefaultSortedOrder
					Case "ASC"
						lsQueryString = lsQueryString & "DESC"
					Case "DESC"
						'do nothing
					Case Else
						lsQueryString = lsQueryString & "ASC"
				End Select
			ElseIf psFieldName = ListOrderField Then
				lsQueryString = lsQueryString & "DESC"
			Else
				lsQueryString = lsQueryString & "ASC"
			End If
		Else
			lsQueryString = lsQueryString & "ASC"
		End If
		
		SortQuery = "?" & lsQueryString
	End Function
' ------------------------------------------------------------------------------
	Public Sub AddSearchedField(ByRef psFieldName)
		
		' Appends another field to memory to be searched.  (Not returned)
		
'		ReDim Preserve msSearchFieldAry(mnSearchFieldCount)
'		msSearchFieldAry(mnSearchFieldCount) = psFieldName

		msSearchFields = msSearchFields & ", " & psFieldName
		mnSearchFieldCount = mnSearchFieldCount + 1
		
	End Sub ' AddSearchedField(ByRef psFieldName)
' ------------------------------------------------------------------------------
	Public Sub AddReturnedField(ByRef psFieldName, ByRef psCaption)
	
		' Appends another field to memory to be returned on the list

		' Field Types: String, Boolean, Number, Other, Date

'		ReDim Preserve msReturnFieldAry(mnReturnFieldCount)
'		ReDim Preserve msReturnFieldCaptionAry(mnReturnFieldCount)
		msReturnFields = msReturnFields & ", " & psFieldName
		msReturnFieldCaptions = msReturnFieldCaptions & ", " & psFieldName

'		msReturnFieldCaptionAry(mnReturnFieldCount) = psCaption
'		msReturnFieldAry(mnReturnFieldCount) = psFieldName

		mnReturnFieldCount = mnReturnFieldCount + 1
		
	End Sub ' AddReturnedField(ByRef psCaption, ByRef psFieldName)
' ------------------------------------------------------------------------------
	Public Function GetTable(ByRef poDatabase)
		
		Dim lsSQL				' Structured Query Language
		Dim lvDataAry			' Data returned from database
		Dim lsTable				' Table of data to be returned

		' Validate that we have needed data and objects
		If Not Validate(poDatabase) Then Exit Function		

		' Setup arrays
		If Not msSearchFields = "" Then
			msSearchFieldAry = Split(Mid(msSearchFields, 3), ", ")
		End If
		If Not msReturnFields = "" Then
			msReturnFieldAry = Split(Mid(msReturnFields, 3), ", ")
			msReturnFieldCaptionAry = Split(Mid(msReturnFieldCaptions, 3), ", ")
		End If	

		' Build SQL to query database
		lsSQL = GetSQL(poDatabase)
		
		' Request page data
		poDatabase.AbsolutePage = mnPage
		poDatabase.PageSize = PageSize
		Call poDatabase.SetData(lsSQL, lvDataAry)
		mnLastPage = poDatabase.PageCount
		mnRecordCount = poDatabase.RecordsAffected
		
		
		' Build the table
		lsTable = _
			"<TABLE Class=""Datalist"">" & _
			GetHeader() & _
			"<FORM name=""DataList"">"
		
		If Not PrimaryKey = "" Then
			lsTable = lsTable & "<INPUT type=""hidden"" name=""pKeyName"" value=""" & PrimaryKey & """>"
		End If

		lsTable = lsTable & _
			GetData(lvDataAry) & _
			"</FORM>" & _
			GetFooter() & _
			"</TABLE>"
		
		Response.Write lsTable
		
	End Function ' GetTable(ByRef poDatabase)
' ------------------------------------------------------------------------------
	Private Function GetHeader()
		Dim lsHTML
		Dim lnIndex
		Dim lnCount
		Dim lsFieldname
		Dim lsCaption
		Dim lnColSpan

		lnColSpan = mnReturnFieldCount
		If Not PrimaryKey = "" Then lnColSpan = lnColSpan + 1
		If ShowTools() Then lnColSpan = lnColSpan + 1

		lsHTML = lsHTML & "<TR class=""Header"">"
		lsHTML = lsHTML & "<TD colSpan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "Found " & mnRecordCount & " Records."
		lsHTML = lsHTML & " Displaying Page " & mnPage & " of " & mnLastPage & "."
		lsHTML = lsHTML & "</TD>"
		lsHTML = lsHTML & "</TR>"

		lsHTML = lsHTML & "<TR class=""Header"">"
		
		If Not PrimaryKey = "" Then
			lsHTML = lsHTML & "<TD>&nbsp;</TD>"
		End If
		
		If AllowSorting Then
			
			lnCount = mnReturnFieldCount - 1
			For lnIndex = 0 To lnCount
				lsFieldName = msReturnFieldAry(lnIndex)
				lsCaption = msReturnFieldCaptionAry(lnIndex)
				
				lsHTML = lsHTML & "<TD><A href=""" & SortQuery(lsFieldName) & """>" & lsCaption & "</A>"
				
				lsHTML = lsHTML & "<IMG src=""" & ImageDirectory
				
				If lsFieldName = msSortedField Then
					Select Case msSortedOrder
						Case "ASC"
							lsHTML = lsHTML & "sort_ascending.gif"
						Case "DESC"
							lsHTML = lsHTML & "sort_descending.gif"
						Case Else
							lsHTML = lsHTML & "sort_none.gif"
					End Select
				ElseIf lsFieldName = DefaultSortedField Then
					Select Case DefaultSortedOrder
						Case "ASC"
							lsHTML = lsHTML & "sort_ascending.gif"
						Case "DESC"
							lsHTML = lsHTML & "sort_descending.gif"
						Case Else
							lsHTML = lsHTML & "sort_none.gif"
					End Select
				ElseIf lsFieldName = ListOrderField Then
					lsHTML = lsHTML & "sort_descending.gif"
				Else
					lsHTML = lsHTML & "sort_none.gif"
				End If
				
				lsHTML = lsHTML & """>"

				
				lsHTML = lsHTML & "</TD>"
			Next
		
		Else
			
			lsHTML = "<TD>" & Join(msReturnFieldAry, "</TD><TD>") & "</TD>"

		End If

		If ShowTools() Then
			lsHTML = lsHTML & "<TD>Tools</TD>"
		End If
		
		lsHTML = lsHTML & "</TR>"
		GetHeader = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Function GetData(ByRef pvDataAry)
		
		Dim lnRowIndex
		Dim lnRowCount
		Dim lnColCount
		Dim lnColIndex
		Dim lsHTML
		Dim lsClass
		Dim lvPrimaryValue
		Dim lnMax
		Dim lsData
		
		If IsArray(pvDataAry) Then
			lnRowCount = UBound(pvDataAry, 2)
		Else
			lnRowCount = -1
		End If
		
		lnMax = PageSize - 1
		lnColCount = mnReturnFieldCount - 1
		
		For lnRowIndex = 0 To lnMax
			If lnRowIndex Mod 2 = 0 Then lsClass = "On" Else lsClass = "Off"
			lsHTML = lsHTML & "<TR class=""" & lsClass & """>"
			
			If Not PrimaryKey = "" Then
				
				If lnRowIndex <= lnRowCount Then
					lvPrimaryValue = pvDataAry(mnReturnFieldCount, lnRowIndex)

					lsHTML = lsHTML & "<TD>"
					lsHTML = lsHTML & "<INPUT" & _
						" type=""radio""" & _
						" class=""Radio""" & _
						" name=""pKey""" & _
						" value=""" & lvPrimaryValue & """"
				
					If lnRowIndex = 0 Then lsHTML = lsHTML & " checked"
				
					lsHTML = lsHTML & "></TD>"
				Else
					lsHTML = lsHTML & "<TD>&nbsp;</TD>"
				End If
			End If
			For lnColIndex = 0 To lnColCount
				lsHTML = lsHTML & "<TD>"
				If lnRowIndex <= lnRowCount Then
					lsData = pvDataAry(lnColIndex, lnRowIndex)
					If VarType(lsData) = vbNull Then
						lsHTML = lsHTML & "&nbsp;"
					Else
						If Len(lsData) > 255 Then
							lsData = lsData & Left(lsData, 255) & " ..."
						End If
						lsData = Server.HTMLEncode(lsData)
						lsHTML = lsHTML & lsData
					End If
				Else
					lsHTML = lsHTML & "&nbsp;"
				End If
				lsHTML = lsHTML & "</TD>"
			Next
			
			If lnRowIndex = 0 Then
				If ShowTools() Then
					lsHTML = lsHTML & "<TD class=""Tools"" rowSpan=""" & PageSize & """>"
					lsHTML = lsHTML & Tools()
					lsHTML = lsHTML & "</TD>"
				End If
			End If
			
			lsHTML = lsHTML & "</TR>"
		Next
		
		GetData = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Sub AppendLink(ByRef psContent, ByRef psLink, ByRef psDelimiter)
		If psContent = "" Then 
			psContent = psLink
		Else
			psContent = psContent & psDelimiter & psLink
		End If
	End Sub
' ------------------------------------------------------------------------------
	Private Function ShowTools()
		If urlAdd = "" And PrimaryKey = "" Then Exit Function

		If urlAdd = "" And urlEdit = "" And urlDelete = "" And urlInsert = "" _
		And urlMoveUp = "" And urlMoveDown = "" Then Exit Function

		ShowTools = True
	End Function
' ------------------------------------------------------------------------------
	Private Function Tools()
		Dim lsLinks
		If Not urlAdd = "" Then
			AppendLink lsLinks, "<A href=""" & urlAdd & """>Add</A>", "<BR>"
		End If
		
		If Not PrimaryKey = "" Then
			If Not urlEdit = "" Then
				AppendLink lsLinks, "<A href=""javascript:datalist_goto('" & urlEdit & "')"">Edit</A>", "<BR>"
			End If

			If Not urlDelete = "" Then
				AppendLink lsLinks, "<A href=""javascript:datalist_goto('" & urlDelete & "')"">Delete</A>", "<BR>"
			End If
		
			If Not ListOrderField = "" Then
				If Not urlInsert = "" Then
					AppendLink lsLinks, "<A href=""javascript:datalist_goto('" & urlInsert & "')"">Insert</A>", "<BR>"
				End If

				If Not urlMoveUp = "" Then
					AppendLink lsLinks, "<A href=""javascript:datalist_goto('" & urlMoveUp & "')"">Move Up</A>", "<BR>"
				End If
				If Not urlMoveDown = "" Then
					AppendLink lsLinks, "<A href=""javascript:datalist_goto('" & urlMoveDown & "')"">Move Down</A>", "<BR>"
				End If
			End If
		End If
		
		Tools = lsLinks
	End Function
' ------------------------------------------------------------------------------
	Private Function PageQuery(ByRef pnPage)
		Dim lsQuery
		
		lsQuery = Request.QueryString
		lsQuery = Replace(lsQuery, "Page=" & mnPage, "")
		
		If pnPage < 1 Then
			pnPage = 1
		ElseIf pnPage > mnLastPage Then
			pnPage = mnLastPage
		End If
		
		If lsQuery = "" Then
			lsQuery = "Page=" & pnPage
		ElseIf Right(lsQuery, 1) = "&" Then
			lsQuery = lsQuery & "Page=" & pnPage
		Else
			lsQuery = lsQuery & "&Page=" & pnPage
		End If
		
		PageQuery = "?" & lsQuery
	End Function
' ------------------------------------------------------------------------------
	Private Function GetFooter()
		Dim lsHTML
		Dim lnColSpan
		Dim lsLinks
		Dim lsPageJump
		Dim lnPage
		Dim lsQueryString
		
		lnColSpan = mnReturnFieldCount
		If Not PrimaryKey = "" Then lnColSpan = lnColSpan + 1
		If ShowTools() Then lnColSpan = lnColSpan + 1
		
		
		If mnLastPage >= 2 Then
			' Create the links.
		
			' First
			AppendLink lsLinks, "<A href=""" & PageQuery(1) & """>First</A>", " | "

			' Previous
			AppendLink lsLinks, "<A href=""" & PageQuery(mnPage - 1) & """>Previous</A>", " | "

			If mnLastPage > 2 Then
				lsQueryString = Request.QueryString
				lsQueryString = Replace(lsQueryString, "Page=" & mnPage, "")
				If lsQueryString = "" Then
					lsQueryString = lsQueryString & "Page="
				ElseIf Right(lsQueryString, 1) = "&" Then
					lsQueryString = lsQueryString & "Page="
				Else
					lsQueryString = lsQueryString & "&Page="
				End If
				lsQueryString = "?" & lsQueryString
				lsPageJump = "<SELECT onChange=""window.location.href = '" & lsQueryString & "' + this[this.selectedIndex].value"">"
				For lnPage = 1 To mnLastPage
					If lnPage = mnPage Then
						lsPageJump = lsPageJump & "<OPTION value=""" & lnPage & """ selected>Page " & lnPage & " of " & mnLastPage & "</OPTION>"
					Else
						lsPageJump = lsPageJump & "<OPTION value=""" & lnPage & """>Page " & lnPage & "</OPTION>"
					End If
				Next
				lsPageJump = lsPageJump & "</SELECT>"
				AppendLink lsLinks, lsPageJump, " | "
			End If
			' Next
			AppendLink lsLinks, "<A href=""" & PageQuery(mnPage + 1) & """>Next</A>", " | "

			' Last
			AppendLink lsLinks, "<A href=""" & PageQuery(mnLastPage) & """>Last</A>", " | "
		
			lsHTML = lsHTML & "<FORM>"
			lsHTML = lsHTML & "<TR class=""Footer"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & lsLinks
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"
			lsHTML = lsHTML & "</FORM>"
		
		End If
		' Do Search
		lsLinks = ""
		lsHTML = lsHTML & "<FORM>"
		lsHTML = lsHTML & "<TR class=""Footer"">"
		lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "<INPUT type=""text"" name=""Query"" value=""" & Server.HTMLEncode(msQuery) & """ size=""20"">"
		lsHTML = lsHTML & " <INPUT type=""submit"" value=""Search"">"
		lsHTML = lsHTML & "</TD>"
		lsHTML = lsHTML & "</TR>"
		lsHTML = lsHTML & "</FORM>"
		
		GetFooter = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Function GetSQL(ByRef poDatabase)
		
		' Returns the SQL to execute against the database.
		
		Dim lsSQL		' Structured Query Language
		Dim lsWhere		' WHERE Clause
		Dim lsSearch	' Users Query translation
		
		' Query for fields returned and the table name
		lsSQL = "SELECT " & Join(msReturnFieldAry, ", ")
		
		If Not PrimaryKey = "" Then lsSQL = lsSQL & ", " & PrimaryKey
		
		lsSQL = lsSQL & " FROM " & TableName

		' Determine if a where clause is needed.
		If Not Arguments = "" Then
			
			' Setup arguments as the where clause
			lsWhere = Arguments
		
		End If ' Not Arguments = ""
		
		lsSearch = poDatabase.QueryFields(msSearchFieldAry, msQuery)
		
		' If a query was built
		If Not lsSearch = "" Then

			' If arguments are present, then append "AND"
			If Not lsWhere = "" Then lsWhere = lsWhere & " AND "

			' Append Search Query to SQL
			lsWhere = lsWhere & lsSearch
		End If
		
		' If arguments have been defined to search against
		If Not lsWhere = "" Then
			
			' Append where clause and arguments to SQL
			lsSQL = lsSQL & " WHERE " & lsWhere
		
		End If ' Not lsWhere = ""
		
		' Determine if query is ordered
		If Not(msSortedField = "" Or msSortedOrder = "") Then
			
			lsSQL = lsSQL & " ORDER BY " & msSortedField & " " & msSortedOrder
			
		ElseIf Not(DefaultSortedField = "" Or DefaultSortedOrder = "") Then
			
			lsSQL = lsSQL & " ORDER BY " & DefaultSortedField & " " & DefaultSortedOrder

		ElseIf Not ListOrderField = "" Then

			lsSQL = lsSQL & " ORDER BY " & ListOrderField & " ASC"

		End If ' Not(msSortedField = "" Or msSortedOrder = "")
		
		GetSQL = lsSQL

	End Function ' GetSQL(ByRef poDatabase)
' ------------------------------------------------------------------------------
	Private Function Validate(ByRef poDatabase)
	
		' If fields have not been setup to be returned, quit
		If Not IsObject(poDatabase) Then 
			Response.Write("<FONT color=red>Object expected.</FONT>")
			Exit Function
		End If

		On Error Resume Next
		If Not poDatabase.ClassName = "LewisMoten.clsDatabase" Then
			Response.Write("<FONT color=red>Database Class object expected.</FONT>")
			Exit Function
		End If
		If Err Then
			Err.Clear
			Response.Write("<FONT color=red>Database Class object expected.</FONT>")
			Exit Function
		End If
		On Error Goto 0

		' If fields have not been setup to be returned, quit
		If mnReturnFieldCount = 0 Then 
			Response.Write("<FONT color=red>Fields have not been requested.</FONT>")
			Exit Function
		End If
		
		' See if a table has not been setup to be queried for data
		If TableName = "" Then
			Response.Write("<FONT color=red>A table has not been defined to query against.</FONT>")
			Exit Function
		End If
		
		Validate = True

	End Function ' Validate(ByRef poDatabase)
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>
<SCRIPT id="Datalist.Scripts">
	<!-- // hide from old browsers (this still needed today?)
	function datalist_goto(psURL)
	{
		var loForm = new Object(document.DataList)
		if(loForm.pKeyName)
		{
			lsName = loForm.pKeyName.value;
			for(var i=0;i<loForm.pKey.length;i++)
			{
				if(loForm.pKey[i].checked)
				{
					window.location.href = psURL + '?' + lsName + '=' + loForm.pKey[i].value;
					return
				}
			}
		}
	}
	// -->
</SCRIPT>