<%
'———————————————————————————————————————————————————————————————————————————————
' Global Declarations
Dim gbShowResultInfo            ' Is the result information shown?
Dim gbShowSearch                ' Is Search form is displayed?

Dim gsArguments                    ' Arguments in SQL string WHERE clause
Dim gsDisplayFieldsArray        ' Fields displayed to user
Dim gsLinkField                    ' fieldname used to link records to other pages.
Dim gsPrimaryKey                ' Unique Identifier of returned data (primary key)
Dim gsRecordSource                ' Table(s) to select records from
Dim gsRetainFieldsArray            ' Seperate query string values to retain.
Dim gsSearchFieldsArray            ' Fields searched when queried.
Dim gsURL_Add                    ' URL to add a new record.
Dim gsURL_Delete                ' URL to delete selected records.
Dim gsURL_Link                    ' URL to link specific records for editing/viewing
Dim gsURL_ThisPage                ' URL for this page.
Dim gsCaptionsAry                ' Captions displayed in each column header

Dim sCellBgColorArray                                ' Background Color
Dim sCellFgColorArray                                ' Foreground Color
Dim sCellFgColorLinkArray                            ' Link Color
Dim sCellClassArray                                    ' Class Names

' Application Settings

' Are we connecting to an SQL Server database?
Const bSQL_DATABASE = False

' How many records are shown on each page.
Const nRECORDS_PER_PAGE = 10

' Text GUI - General User Interface

' Font face/family
Const sFONT_FACE = "Helvetica,Arial,sans-serif"
Const sFONT_SIZE = 2

' Colors

' Cell Headers
Const sBGCOLOR_CELL_HEADER            = "#cccccc"        ' Background Color
Const sFGCOLOR_CELL_HEADER            = "#444444"        ' Foreground Color
Const sFGCOLOR_CELL_HEADER_LINK        = "#000000"        ' Link Color

' Alternaing Cell styles

' Store alternating Background colors
sCellBgColorArray = Array("#dddddd", "#ffffff")

' Store alternating Foreground colors
sCellFgColorArray = Array("#000000")

' Store alternating Link Colors
sCellFgColorLinkArray = Array("#000000")

' Store alternating Class Names
sCellClassArray = Array("HiLite", "NoLite")

Const sBORDER_COLOR					= "#808080"

Const sBGCOLOR_SEARCH                = "#cccccc"        ' Background Color
Const sFGCOLOR_SEARCH                = "#000000"        ' Foreground Color

' Page Navigation colors
Const sBGCOLOR_NAVIGATION            = "#cccccc"        ' Background Color
Const sFGCOLOR_NAVIGATION            = "#888888"        ' Foreground Color
Const sFGCOLOR_NAVIGATION_LINK        = "#000000"        ' Link Color

' text color for NULL values.
Const sFGCOLOR_NULL                    = "#f8f8f8"        ' Foreground Color

' text color for EMPTY values.
Const sFGCOLOR_EMPTY                = "#f8f8f8"        ' Foreground Color

' Graphic GUI - General User Interface

' Image directory
Const sIMAGE_ROOT                        = "/images/"

' page navigation
Const sIMAGE_NAV_FIRST_PAGE                = "nav_first.gif"
Const sIMAGE_NAV_FIRST_PAGE_DISABLED    = "nav_first_disabled.gif"

' image links for previous page
Const sIMAGE_NAV_PREVIOUS_PAGE            = "nav_previous.gif"
Const sIMAGE_NAV_PREVIOUS_PAGE_DISABLED    = "nav_previous_disabled.gif"

' image links for next page
Const sIMAGE_NAV_NEXT_PAGE                = "nav_next.gif"
Const sIMAGE_NAV_NEXT_PAGE_DISABLED        = "nav_next_disabled.gif"

' image links for last page
Const sIMAGE_NAV_LAST_PAGE                = "nav_last.gif"
Const sIMAGE_NAV_LAST_PAGE_DISABLED        = "nav_last_disabled.gif"

' Sort Order
Const sIMAGE_SORT_DESCENDING            = "sort_descending.gif"
Const sIMAGE_SORT_ASCENDING                = "sort_ascending.gif"
Const sIMAGE_SORT_NONE                    = "sort_none.gif"

' Buttons
Const sIMAGE_BUTTON_DELETE                = "delete.gif"
Const sIMAGE_BUTTON_ADD                    = "add.gif"
Const sIMAGE_BUTTON_SEARCH                = "find.gif"
'———————————————————————————————————————————————————————————————————————————————
Sub Data_GetListData( _
    ByRef s_SQL, _
    ByRef n_RecordsPerPage, _
    ByRef n_AbsolutePage, _
    ByRef n_TotalRecords, _
    ByRef n_MaxPages, _
    ByRef s_FieldNameArray, _
    ByRef v_ListDataArray, _
    ByRef n_LowestRecordNum, _
    ByRef n_HighestRecordNum _
    )

    'Dim lsACTIVE_CONNECTION = "DSN=allowancenet;User=sa;password=808UTD.State"
    
    Dim loField                ' ADODB.Field object (from loRS.Fields)
    Dim loRS                ' ADODB.Recordset object
    Dim lsFieldNames        ' List of field names returned from database
    
    ' Create object to request and retain data.
    Set loRS = Server.CreateObject("ADODB.Recordset")
    
    ' Send SQL String to debug routine: WriteDebug()
    Call WriteDebug("SQL", s_SQL)

    ' Connect to database and request data
    Call loRs.Open( _
        s_SQL, _
        sCONNECTION_STRING, _
        adOpenKeyset, _
        adLockBatchOptimistic, _
        adCmdText)

    ' Store how many records were returned
    n_TotalRecords = loRS.RecordCount
    
    ' If no records were returned
    If n_TotalRecords = 0 Then
    
        ' Release objects from memory
        Set loRS = Nothing

        ' Send message to client
        'Response.Write("<P>No results were found matching the query.</P>")

        ' Exit the routine
        Exit Sub
        
    End If ' n_TotalRecords = 0

    ' Set the page size - amount of records on each page
    loRS.PageSize = n_RecordsPerPage
    
    ' Store how many pages were returned
    n_MaxPages = loRS.PageCount

    ' If the last page is less then the page being viewd
    If n_MaxPages < n_AbsolutePage Then
    
        ' Send error message to client
        Response.Write("<P>You are currently viewing a page that is out " & _
            "of range(1 - " & n_MaxPages & "). Perhaps some records " & _
            "have been deleted.</P>")

        ' Exit the routine
        Exit Sub
        
    End If ' n_MaxPages < lnAbsolutePage

    ' Move through the records to the absolute page
    loRS.AbsolutePage = n_AbsolutePage

    ' Loop through each Field Name returned from the database
    For Each loField In loRS.Fields
    
        ' Add name to list of field names returned delimited by tabs
        lsFieldNames = lsFieldNames & vbTab & loField.Name

    Next ' loField
        
    ' Remove first extra tab from Field Name list
    lsFieldNames = Mid(lsFieldNames, 2)
    
    ' Parse field name list into an array
    s_FieldNameArray = Split(lsFieldNames, vbTab)
    
    ' Store the records displayed on the current page into an array
    v_ListDataArray = loRS.GetRows(nRECORDS_PER_PAGE)', adBookmarkCurrent)

    ' Release Data Objects from memory
    Set loField = Nothing        ' ADODB.Fields
    Set loRS = Nothing            ' ADODB.Recordset

    ' Store the position of the first record on the page
    n_LowestRecordNum = (nRECORDS_PER_PAGE * (n_AbsolutePage -1)) + 1
    
    ' Store the position of the last record on the page
    n_HighestRecordNum = n_LowestRecordNum + UBound(v_ListDataArray, 2)

End Sub
'———————————————————————————————————————————————————————————————————————————————
Function HTML_BuildHeader( _
    ByRef s_FieldNameArray, _
    ByRef n_PrimaryKeyIndex, _
    ByRef n_LinkColumnIndex _
    )
     
    Dim lsResult
    Dim lnIndex
    Dim lsQueryString
    Dim lbRendureCell
    Dim lsSortField
    Dim lsSort
    
    'gsNonAliasedFieldsArray
    
    lsSortField = Request.QueryString("SortField")
    lsSort = Request.QueryString("Sort")
    
    ' Write opening comments
    lsResult = vbCrLf & "<!-- Header: BEGIN -->" & vbCrLf

    ' Begin new row.
    lsResult = lsResult & "<TR>" & vbCrLf

    ' Acquire Querystring to pass existing querystring values (except Sort and Order).
    lsQueryString = NewQueryString(Array("SortField", "Sort"))
    
    
    ' Loop through each field name to find Primary key Index
    For lnIndex = 0 To UBound(s_FieldNameArray, 1)

        ' No - lets compare field name against defined primary key.
        If LCase(s_FieldNameArray(lnIndex)) = LCase(gsPrimaryKey) Then

            ' Store the index of the primary key.
            n_PrimaryKeyIndex = lnIndex

            Exit For ' lnIndex

        End If ' s_FieldNameArray(lnIndex) = gsPrimaryKey
    
    Next
    
    ' Loop through each field name to find Link Column Index
    For lnIndex = 0 To UBound(s_FieldNameArray, 1)

        ' Check name against Linking field name
        If LCase(s_FieldNameArray(lnIndex)) = LCase(gsLinkField) Then

            ' Store the index of the Linking Column
            n_LinkColumnIndex = lnIndex
            
            Exit For ' lnIndex
        
        End If
    
    Next ' lnIndex
    
    ' see if the field to be linked to edit records has been set.
    If n_LinkColumnIndex = -1 Then

        ' No, Mark the primary field as the link column for default.
        n_LinkColumnIndex = n_PrimaryKeyIndex

    End If ' lnLinkColumnIndex = -1

    ' Loop through each field name
    For lnIndex = 0 To UBound(s_FieldNameArray, 1)

        ' Determine if we are working with primary key
        If lnIndex = n_PrimaryKeyIndex Then
        
            ' Determine if another field is to be used for linking.
            If n_LinkColumnIndex = n_PrimaryKeyIndex Then
            
                ' No, the current cell is to be used for linking.
                lbRendureCell = True
            
            Else
            
                ' Yes, flag cell "Not for Renduring"
                lbRendureCell = False
                
            End If ' lnLinkColumnIndex = lnPrimaryKeyIndex
        
        Else
        
            ' Display the cell
            lbRendureCell = True

        End If ' lnIndex = lnPrimaryKeyIndex
        
        ' Check to see if it is ok to rendure the cell.
        If lbRendureCell Then

            ' Begin building a table cell.
            lsResult = lsResult & "<TD bgcolor=""" & _
                sBGCOLOR_CELL_HEADER & """ align=""center"" " & _
                "class=""Cell-Header"">" & vbCrLf
            
            ' Add the text style
            lsResult = lsResult & "<FONT size=""1"" face=""" & _
                sFONT_FACE & """ color=""" & sFGCOLOR_CELL_HEADER & """>" & _
                vbCrLf & "<B>" & vbCrLf
            
            ' Determine if client has sorted this field.
            If lsSortField = s_FieldNameArray(lnIndex) And Not lsSort = "" Then

                ' Yes, determine how the field is to be sorted.
                Select Case UCase(lsSort)

                    ' Ascending order
                    Case "ASC"
                        
                        ' Check to see if an image is to be used for showing the sort.
                        If sIMAGE_SORT_ASCENDING = "" Then
                            
                            ' No - place marker before field name
                            lsResult = lsResult & "[" & vbCrLf
                            
                        End If ' sIMAGE_SORT_ASCENDING = ""
                    
                    ' Descending Order
                    Case "DESC"
                        
                        ' Check to see if an image is to be used for showing the sort.
                        If sIMAGE_SORT_DESCENDING = "" Then
                            
                            ' No - place marker before field name
                            lsResult = lsResult & "[" & vbCrLf
                            
                        End If ' sIMAGE_SORT_DESCENDING = ""
                
                End Select ' UCase(lsSort)
            
            End If ' lsSortField = s_FieldNameArray(lnIndex) And Not lsSort = ""

            ' Determine if field name matches its counterpart in the display fields
            ' (some complex queries can not be sorted)
            If s_FieldNameArray(lnIndex) = gsDisplayFieldsArray(lnIndex) Then
            
                ' Create the link and encode the field name to be passed.
                lsResult = lsResult & "<A href=""" & gsURL_ThisPage & "?" & _
                    lsQueryString & "SortField=" & Server.URLEncode( _
                    s_FieldNameArray(lnIndex)) & "&Sort="

                ' Check to see if the user wants this field sorted.
                If lsSortField = s_FieldNameArray(lnIndex) Then
                    
                    ' Determine the sort order
                    Select Case UCase(lsSort)
                        
                        ' Ascending
                        Case "ASC"
                            
                            ' Write querystring - next order is descending
                            lsResult = lsResult & "DESC"
                        
                        ' Descending
                        Case "DESC"
                            
                            ' Do Nothing (defaults to no sort order)
                        
                        ' None
                        Case Else
                            
                            ' Write querystring - next order is ascending
                            lsResult = lsResult & "ASC"
                            
                    End Select ' UCase(lsSort)
                
                Else

                    ' Set unsorted fields to sort ascending if clicked.
                    lsResult = lsResult & "ASC"

                End If ' lsSortField = s_FieldNameArray(lnIndex)
                
                ' Close the <A> link
                lsResult = lsResult & """ style=""color:" & _
                    sFGCOLOR_CELL_HEADER_LINK & ";"">"
            
            End If ' s_FieldNameArray(lnIndex) = gsDisplayFieldsArray(lnIndex)
            
            ' Determine if we are looking at the primary key field
            If n_PrimaryKeyIndex = lnIndex Then
            
                If IsArray(gsCaptionsAry) Then
                    ' Write the name of the field
                    lsResult = lsResult & gsCaptionsAry(0)
                Else
                    ' Write Field name as "ID"
                    lsResult = lsResult & "ID"
                End If
            Else
                If IsArray(gsCaptionsAry) Then
                    If UBound(gsCaptionsAry, 1) >= lnIndex Then
                        ' Write the name of the field
                        lsResult = lsResult & gsCaptionsAry(lnIndex)
                    Else
                        ' Write the name of the field
                        lsResult = lsResult & s_FieldNameArray(lnIndex)
                    End If
                Else
                    ' Write the name of the field
                    lsResult = lsResult & s_FieldNameArray(lnIndex)
                End If
            End If ' gsPrimaryKey = s_FieldNameArray(lnIndex)
            
            ' Determine if field had been linked for sorting.
            If s_FieldNameArray(lnIndex) = gsDisplayFieldsArray(lnIndex) Then
                
                ' Close the link
                lsResult = lsResult & "</A>"

                ' Determine if user requested sort on this field.
                If lsSortField = s_FieldNameArray(lnIndex) And Not lsSort = "" Then
                
                    ' Determine how the link is sorted.
                    Select Case UCase(lsSort)
                    
                        ' Ascending
                        Case "ASC"
                            
                            ' Check to see if an image is to be used for showing the sort.
                            If sIMAGE_SORT_ASCENDING = "" Then
                                
                                ' No image, show text disply of sort properties.
                                lsResult = lsResult & "&nbsp;" & lsSort & _
                                    "&nbsp;]" & vbCrLf

                            Else
                            
                                ' Add the image showing sort order
                                lsResult = lsResult & "<IMG src=""" & _
                                    sIMAGE_ROOT & sIMAGE_SORT_ASCENDING & _
                                    """ border=""0"" " & _
                                    "alt=""Ascending Order"" " & _
                                    "hspace=""10"" vspace=""2"" " & _
                                    "align=""absmiddle"" />" & vbCrLf
                            
                            End If ' sIMAGE_SORT_ASCENDING = ""
                        
                        ' Descending
                        Case "DESC"
                            
                            ' Check to see if an image is to be used for showing the sort.
                            If sIMAGE_SORT_DESCENDING = "" Then
                            
                                ' No image, show text disply of sort properties.
                                lsResult = lsResult & "&nbsp;" & lsSort & _
                                    "&nbsp;]" & vbCrLf
                            
                            Else
                            
                                ' Add the image showing sort order
                                lsResult = lsResult & "<IMG src=""" & _
                                    sIMAGE_ROOT & sIMAGE_SORT_DESCENDING & _
                                    """ border=""0"" " & _
                                    "alt=""Descending Order"" " & _
                                    "hspace=""10"" vspace=""2"" " & _
                                    "align=""absmiddle"" />" & vbCrLf

                            End If

                        Case Else ' Not sorted

                            ' Check to see if an image is to be used for showing the sort.
                            If Not sIMAGE_SORT_NONE = "" Then
                            
                                ' Add the image showing sort order
                                lsResult = lsResult & "<IMG src=""" & _
                                    sIMAGE_ROOT & sIMAGE_SORT_NONE & _
                                    """ border=""0"" " & _
                                    "alt=""Not Sorted"" " & _
                                    "hspace=""10"" vspace=""2"" " & _
                                    "align=""absmiddle"" />" & vbCrLf
                                
                            End If ' Not sIMAGE_SORT_NONE = ""
                            
                    End Select ' UCase(lsSort)
                
                Else
                
                    ' Check to see if an image is to be used for showing the sort.
                    If Not sIMAGE_SORT_NONE = "" Then
                            
                        ' Add the image showing sort order
                        lsResult = lsResult & "<IMG src=""" & _
                            sIMAGE_ROOT & sIMAGE_SORT_NONE & _
                            """ border=""0"" " & _
                            "alt=""Not Sorted"" " & _
                            "hspace=""10"" vspace=""2"" " & _
                            "align=""absmiddle"" />" & vbCrLf
                                
                    End If ' Not sIMAGE_SORT_NONE = ""
                    
                End If ' lsSortField = s_FieldNameArray(lnIndex) And Not lsSort = ""
                
            End If ' s_FieldNameArray(lnIndex) = gsDisplayFieldsArray(lnIndex)

            ' Close the table cell
            lsResult = lsResult & _
                "</B>" & vbCrLf & _
                "</FONT>" & vbCrLf & _
                "</TD>" & vbCrLf
        
        Else

            ' Field is hidden. do not show the contents.

        End If ' lbRendureCell
        
    Next ' lnIndex
    
    ' Determine if a URL is provided for deleting multiple records.
    If Not gsURL_DELETE = "" Then
    
        ' ( Write delete cell header )
        
        ' Write a comment noting where the html begins for delete header.
        lsResult = lsResult & "<!-- Header - Delete: BEGIN -->" & vbCrLf
        
        ' Open the delete cell
        lsResult = lsResult & "<TD bgcolor=""" & sBGCOLOR_CELL_HEADER & _
            """ class=""Cell-Header"" align=""center"" width=""1%"">" & vbCrLf
        
        ' Add the font style.
         lsResult = lsResult & "<FONT size=""1"" face=""" & sFONT_FACE & _
            """ color=""" & sFGCOLOR_CELL_HEADER & """>" & vbCrLf
        
        ' Add the text "Delete"
        lsResult = lsResult & "<B>Delete<BR />Checked</B>" & vbCrLf
        
        ' Close teh cell
        lsResult = lsResult & "</FONT>" & vbCrLf & "</TD>" & vbCrLf
        
        ' Write closing comments
        lsResult = lsResult & "<!-- Header - Delete: END -->" & vbCrLf
        
    End If ' Not gsURL_DELETE = ""
    
    ' Close the header row.
    lsResult = lsResult & "</TR>" & vbCrLf

    ' Write closing header comments
    lsResult = lsResult & "<!-- Header: END -->" & vbCrLf
    
    ' Return the results.
    HTML_BuildHeader = lsResult
    
End Function ' HTML_BuildHeader
'———————————————————————————————————————————————————————————————————————————————
Function HTML_BuildInformation( _
    ByRef n_TotalRecords, _
    ByRef n_Page, _
    ByRef n_LastPage, _
    ByRef n_LowestRecordNum, _
    ByRef n_HighestRecordNum _
    )
    
    Dim lsQuery                ' Text client is searching for.
    Dim lsResults            ' HTML code to be passed back.
    
    lsQuery = Request.QueryString("Query")
    
    ' Open the code to display the data
    lsResults = _
        "<!-- List Results Details: BEGIN -->" & vbCrLf & _
        "<CENTER>" & vbCrLf & _
        "<TABLE width=""80%"" class=""List-Results"">" & vbCrLf & _
        "<TR>" & vbCrLf & _
        "<TD>" & vbCrLf & _
        "<FONT size=""2"" face=""" & sFONT_FACE & """>" & vbCrLf & _
        "<P>" & vbCrLf
    
    ' Display the number of records found
    lsResults = lsResults & _
        "A total of " & vbCrLf & _
        "<NOBR><B>" & n_TotalRecords & "</B> records</NOBR>" & _
        " have been found"

    ' Determine if the user tried to search for records.
    If Not lsQuery = "" Then
        lsResults = lsResults & " matching your query"
    End If ' Not lsQuery = ""

    ' Finnish the sentance.
    lsResults = lsResults & " in the database." & vbCrLf
    
    ' Display wich page is being viewed, how many pages were found, and how
    ' many records are displayed on each page.
    lsResults = lsResults & _
        " You are now viewing " & _
        "<NOBR>page " & n_Page & " of " & n_LastPage & "</NOBR>" & vbCrLf & _
        " with " & nRECORDS_PER_PAGE & " records on each page."

    ' Display the results ordinal position range.
    lsResults = lsResults & _
        " Now showing results " & vbCrLf & _
        "<NOBR>" & vbCrLf & _
        "from " & _
        "<B>" & n_LowestRecordNum & "</B>" & vbCrLf & _
        " to " & vbCrLf & _
        "<B>" & n_HighestRecordNum & "</B>" & vbCrLf & _
        "</NOBR>"
    
    ' Close the code to display the data
    lsResults = lsResults & _
        "</P>" & vbCrLf & _
        "</FONT>" & vbCrLf & _
        "</TD>" & vbCrLf & _
        "</TR>" & vbCrLf & _
        "</TABLE>" & vbCrLf & _
        "</CENTER>" & vbCrLf & _
        "<!-- List Results Details:END -->" & vbCrLf

    ' Return the HTML code
    HTML_BuildInformation = lsResults

End Function ' HTML_BuildInformation
'———————————————————————————————————————————————————————————————————————————————
Private Function HTML_BuildNavigation(ByRef n_AbsolutePage, ByRef n_PageCount) ' As String
    
    ' n_AbsolutePage        ' Current page that user is looking at.
    ' n_PageCount            ' Maximum pages that can be viewed.
    
    Dim lnIndex                ' Loop Counter

    Dim lsCellClose            ' Closing table cell HTML code
    Dim lsCellOpen            ' Opening table cell HTML code
    Dim lsFieldName            ' Name of current field in a collection
    Dim lsQueryString        ' Querystring to write within link.
    Dim lsResult            ' text returned from procedure
    Dim lbIsFirstPage        ' Is user on the first page?
    Dim lbIsLastPage        ' Is user on the last page?
    
    ' Determine if there are at least 2 pages.
    If n_PageCount < 2 Then

        ' Get out of this routine to return no navigation.
        Exit Function
        

    End If ' n_PageCount < 2
    
    '? do we need this?
    '? do we need this?
    
    ' Get all variables in current query string except "Page"
    lsQueryString = NewQueryString(Array("Page"))
    
    '? do we need this?
    '? do we need this?
    
    
    ' Define the opening table cell code.
    lsCellOpen = vbCrLf & "<TD align=""center"" bgcolor=""" & sBGCOLOR_NAVIGATION & _
        """ class=""lstRecords-Navigation"" width=""20%"">" & vbCrLf
    
    ' Add the font style.
    lsCellOpen = lsCellOpen & "<FONT size=""2"" face=""" & sFONT_FACE & _
        """ color=""" & sFGCOLOR_NAVIGATION & """>" & vbCrLf
    
    ' Define cell closing HTML
    lsCellClose = "</FONT>" & vbCrLf & "</TD>" & vbCrLf
    
    ' Begin the results with an opening comment.
    lsResult = vbCrLf & "<!-- List Navigation: BEGIN -->" & vbCrLf
    
    ' Open the navigation table and row.
    lsResult = lsResult & "<TABLE width=""100%"" border=""0"" " & _
        "cellpadding=""0"" cellspacing=""0"">" & vbCrLf & "<TR>" & vbCrLf
    
    
    ' Determine if we are on the first page
    lbIsFirstPage = (n_AbsolutePage = 1)

    ' Determine if user is viewing last page.
    lbIsLastPage = (n_AbsolutePage = n_PageCount)
    
    ' Write Link to first page
    lsResult = lsResult & lsCellOpen & HTML_NavLink( _
        lbIsFirstPage, _
        0, _
        "|&lt; First Page", _
        sIMAGE_NAV_FIRST_PAGE, _
        sIMAGE_NAV_FIRST_PAGE_DISABLED _
        ) & lsCellClose
        
    ' Write Link to previous page
    lsResult = lsResult & lsCellOpen & HTML_NavLink( _
        lbIsFirstPage, _
        n_AbsolutePage - 1, _
        "&lt;&lt; Previous Page", _
        sIMAGE_NAV_PREVIOUS_PAGE, _
        sIMAGE_NAV_PREVIOUS_PAGE_DISABLED _
        ) & lsCellClose

    'Build a form to allow jumps to any page with a select box.
    lsResult = lsResult & "<FORM action=""" & gsURL_ThisPage & _
        """ method=""get"" id=""NavigationForm"" name=""NavigationForm"">" & _
        vbCrLf

    ' pass querystring data within the form.
    For Each lsFieldName In Request.QueryString

        'Don't write the page - thats up to the user to choose
        If Not LCase(lsFieldName) = "page" Then

            ' Write a hidden input field
            lsResult = lsResult & "<INPUT type=""hidden"" name=""" & _
                lsFieldName & """ value=""" & _
                Request.QueryString(lsFieldName) & """ />" & vbCrLf

        End If ' Not LCase(lsFieldName) = "page"

    Next ' lsFieldName

    ' Add a select box to hold valid pages to jump to
    lsResult = lsResult & lsCellOpen & "<SELECT onChange=""if(this.selectedIndex!=" & (n_AbsolutePage - 1) & "){window.document.NavigationForm.submit();}"" name=""Page"">" & vbCrLf
    
    ' Loop through all available page numbers.
    For lnIndex = 1 To n_PageCount
    
        ' Begin to write option statement.
        lsResult = lsResult & "<OPTION value=""" & lnIndex & """"
        
        ' Determine if user is viewing current page.
        If lnIndex = n_AbsolutePage Then
        
            ' Yes, mark as selected.
            lsResult = lsResult & " selected=""True"">"
            
            ' Write details about the current page
            lsResult = lsResult & "Page " & lnIndex & " of " & n_PageCount
            
        Else
        
            ' No, end the select tag and write the page number.
            lsResult = lsResult & ">" & "Page " & lnIndex
        
        End If ' lnIndex = n_AbsolutePage
        
        ' Close option statment.
        lsResult = lsResult & "</OPTION>" & vbCrLf
        
    Next ' lnIndex
    
    ' Close select tag
    lsResult = lsResult & "</SELECT>"
    
    ' Add a submit button.
    lsResult = lsResult & "<INPUT type=""submit"" value=""GO"" id=""Submit"" name=""Submit"" />" & vbCrLf
    
    ' Close the cell
    lsResult = lsResult & lsCellClose
    
    ' Close the form
    lsResult = lsResult & "</FORM>" & vbCrLf
    
    ' No, Write Link to next page
    lsResult = lsResult & lsCellOpen & HTML_NavLink( _
        lbIsLastPage, _
        n_AbsolutePage + 1, _
        "Next Page &gt;&gt;", _
        sIMAGE_NAV_NEXT_PAGE, _
        sIMAGE_NAV_NEXT_PAGE_DISABLED _
        ) & lsCellClose
            
    ' Write Link to last page
    lsResult = lsResult & lsCellOpen & HTML_NavLink( _
        lbIsLastPage, _
        n_PageCount, _
        "Last Page &gt;|", _
        sIMAGE_NAV_LAST_PAGE, _
        sIMAGE_NAV_LAST_PAGE_DISABLED _
        ) & lsCellClose
    
    ' Close Navigation row
    lsResult = lsResult & "</TR>" & vbCrLf
    
    ' Close Navigation table
    lsResult = lsResult & "</TABLE>" & vbCrLf
    
    ' Write Closing Comments.
    lsResult = lsResult & vbCrLf & "<!-- List Navigation: END -->" & vbCrLf
    
    ' return the results
    HTML_BuildNavigation = lsResult
    
End Function ' HTML_BuildNavigation
'———————————————————————————————————————————————————————————————————————————————
Function HTML_Form_Search( _
    ByRef s_Keywords, _
    ByRef s_MatchWords, _
    ByRef s_MatchType, _
    ByRef s_RetainerArray, _
    ByRef s_SearchImage _
    )

    ' Function HTML_Search()
    '
    ' This routine outputs HTML code to build a search form and populate
    ' it depending on the information supplied by the calling routine.
    '
    ' Incomming Variables
        ' Keywords to look for.
        ' Match records ([With]|Without) keywords
        ' Match ([Any]|All|Phrase) Keywords
        ' List of unrelated field names to retain from querystring
        ' Image to display as search button
        '
    
    ' Declare local variables
    Dim lsResult            ' text that is returned form the procedure.
    Dim lsCellOpen            ' Opening table cell HTML code
    Dim lsCellClose            ' Closing table cell HTML code
    Dim lsFieldName            ' Name of current field in QueryString
    
    ' Store HTML to define an opening table cell.
    lsCellOpen = vbCrLf & "<TD align=""center"" bgcolor=""" & _
        sBGCOLOR_SEARCH & """>" & vbCrLf & "<FONT size=""2"" face=""" & _
        sFONT_FACE & """ color=""" & sFGCOLOR_SEARCH & """>" & vbCrLf
    
    ' Store HTML to define a closing table cell
    lsCellClose = vbCrLf & "</FONT>" & vbCrLf & "</TD>" & vbCrLf
    
    ' Store opening comment in results
    lsResult = "<!-- Search Form: BEGIN -->" & vbCrLf

    ' Add opening table to results
    lsResult = lsResult & _
        "<TABLE width=""100%"" bgcolor=""" & sBGCOLOR_SEARCH & _
        """ cellspacing=""0"" border=""0"" class=""lstRecords-Search"">" & _
        vbCrLf
    
    ' Add opening form tag to results
    lsResult = lsResult & "<FORM action=""" & gsURL_ThisPage & """ " & _
        "method=""get"" id=""Search"" name=""Search"">" & vbCrLf

    ' Loop through a list of field names to retain from routine input
    For Each lsFieldName In s_RetainerArray
        
        ' Add a hidden tag to store the retained data
        lsResult = lsResult & _
            "<INPUT type=""hidden"" name=""" & lsFieldName & """ value=""" & _
            Request.QueryString(lsFieldName) & """ />" & vbCrLf

    Next ' lsFieldName
    
    ' Add opening table row <TR> to results
    lsResult = lsResult & "<TR>" & vbCrLf
    
    ' Add opening table cell to results
    lsResult = lsResult & lsCellOpen
    
    ' Add caption "Search" to results
    lsResult = lsResult & "<B>Search:</B>"
    
    ' Add closing table cell to results
    lsResult = lsResult & lsCellClose
    
    ' Add opening table cell to results
    lsResult = lsResult & lsCellOpen

    ' Add input (text) for keywords to results
    lsResult = lsResult & "<INPUT type=""text"" name=""Query"" " & _
        "size=""6"" style=""width:90px;"" value="""
        
    ' Add (Keywords) routine input for default value for keywords to results
    lsResult = lsResult & s_Keywords
    
    ' Add closing input (text) for keywords to results
    lsResult = lsResult & """ />" & vbCrLf
    
    ' Add closing table cell to results
    lsResult = lsResult & lsCellClose
    
    ' Add opening table cell to results
    lsResult = lsResult & lsCellOpen
    
    ' Add opening <SELECT id=select1 name=select1> tag (MatchWords) to results
    lsResult = lsResult & "<SELECT name=""Match1"">" & vbCrLf
    
    ' ADD option to match fields (with) words to results
    lsResult = lsResult & "<OPTION value=""With"""
    
    ' IF NOT routine input states to look for data (without) words
    If Not LCase(s_MatchWords) = "without" Then
        
        ' Add (with) as selected TO results
        lsResult = lsResult & " selected=""True"""
    
    End If ' Not LCase(s_MatchWords) = "without"
    
    ' ADD closing </OPTION> tag (with) TO results
    lsResult = lsResult & ">With</OPTION>" & vbCrLf
    
    ' ADD opening <OPTION> tag (without) TO results
    lsResult = lsResult & "<OPTION value=""Without"""

    ' IF routine input states to look for data (without) words
    If LCase(s_MatchWords) = "without" Then
    
        ' Add (without) as selected TO results
        lsResult = lsResult & " selected=""True"""
        
    End If ' LCase(s_MatchWords) = "without"
    
    ' ADD closing </OPTION> tag (without) TO results
    lsResult = lsResult & ">Without</OPTION>" & vbCrLf
    
    ' ADD closing </SELECT> tag (MatchWords) TO results
     lsResult = lsResult & "</SELECT>" & vbCrLf
    
    ' Add opening <SELECT id=select1 name=select1> tag (MatchType) to results
    lsResult = lsResult & "<SELECT name=""Match2"">" & vbCrLf
    
    ' Add opening <OPTION> tag (Any) to results
    lsResult = lsResult & "<OPTION value=""Any"""
    
    ' IF NOT routine input states to match (all words) OR (phrase)
    If Not(LCase(s_MatchType) = "all" Or LCase(s_MatchType) = "phrase") Then
    
        ' ADD (Match ANY Word) as selected TO results
        lsResult = lsResult & " selected=""True"""
        
    End If ' Not(LCase(s_MatchType) = "all" Or LCase(s_MatchType) = "phrase")
    
    ' ADD closing </OPTION> tag (Match ANY word) TO results
    lsResult = lsResult & ">Any Word</OPTION>" & vbCrLf
        
    ' ADD opening <OPTION> to (Match ALL words) TO results
    lsResult = lsResult & "<OPTION value=""All"""
    
    ' IF routine input states to match all words
    If LCase(s_MatchType) = "all" Then
        
        ' ADD (Match ALL Words) as selected to results
        lsResult = lsResult & " selected=""True"""
    
    End If ' LCase(lsMatch2) = "all"
    
    ' ADD closing </OPTION> to (Match ALL Words) TO results
    lsResult = lsResult & ">All Words</OPTION>" & vbCrLf
    
    ' ADD option (Match phrase) TO results
    lsResult = lsResult & "<OPTION value=""Phrase"""
    
    ' IF routine input states (match by phrase)
    If LCase(s_MatchType) = "phrase" Then
        
        ' ADD (Match PHRASE) as selected to results
        lsResult = lsResult & " selected=""True"""
    
    End If ' LCase(s_MatchType) = "phrase"
    
    ' ADD closing </OPTION> to (Match PHRASE) TO results
    lsResult = lsResult & ">Phrase</OPTION>" & vbCrLf
    
    ' ADD closing </SELECT> tag (type of match) TO results
    lsResult = lsResult & "</SELECT>" & vbCrLf
    
    ' ADD closing table cell TO results
    lsResult = lsResult & lsCellClose
    
    ' Add opening table cell for the search button to results
    lsResult = lsResult & lsCellOpen
    
    ' If graphic was not provided for submiting results
    If s_SearchImage = "" Then
    
        ' ADD <INPUT id=text1 name=text1> (submit) to results
        lsResult = lsResult & "<INPUT type=""submit"" value=""Search"">"
    
    ' Else a graphic was provided
    Else
    
        ' Add <INPUT id=text1 name=text1> (image) to results
        lsResult = lsResult & "<INPUT type=""image"" " & _
            "value=""Search"" src=""" & sIMAGE_ROOT & s_SearchImage & _
            """ alt=""Search"" " & "border=""0"" / id=1 name=1>"
    
    End If 'sIMAGE_BUTTON_SEARCH = ""
    
    ' Add closing table cell for search button to results
    lsResult = lsResult & lsCellClose
    
    ' Add closing table row
    lsResult = lsResult & "</TR>" & vbCrLf
    
    ' Add closing </FORM>
    lsResult = lsResult & "</FORM>" & vbCrLf
    
    ' Add closing </TABLE>
    lsResult = lsResult & "</TABLE>" & vbCrLf

    ' Add closing comment
    lsResult = lsResult & vbCrLf & "<!-- Search Form: END -->" & vbCrLf
    
    ' output the results to the calling routine
    HTML_Form_Search = lsResult
    
End Function ' HTML_Form_Search
'———————————————————————————————————————————————————————————————————————————————
Private Function HTML_NavLink( _
    ByRef b_IsDisabled, _
    ByRef n_Page, _
    ByRef s_Caption, _
    ByRef s_Image, _
    ByRef s_Image_Disabled _
    )
    
    ' b_IsDisabled            ' Is the link going to show up as disabled?
    ' n_Page                ' page that link will point to
    ' s_Caption                ' Text to be displayed in the link (or Alt )
    ' s_Image                ' Image to be displayed within link.
    ' s_Image_Disabled        ' Disabled version of the image
    
    Dim lsQueryString
    Dim lsResult
    
    ' Check to see if the navigation is disabled.
    If b_IsDisabled Then
    
        ' Yes, Was a disabled image specified?
        If s_Image_Disabled = "" Then
            
            ' No, was a regular image specified?
            If s_Image = "" Then
                
                ' No, create disabled text
                lsResult = "<U>" & s_Caption & "</U>" & vbCrLf
            
            Else
            
                ' Yes, create regular image
                lsResult = "<IMG src=""" & sIMAGE_ROOT & s_Image & """ border=""0"" alt=""" & _
                    s_Caption & """ hspace=""10"" vspace=""2"" />"

            End If ' s_Image = ""

        Else
    
            ' Yes, create disabled image
            lsResult = "<IMG src=""" & sIMAGE_ROOT & s_Image_Disabled & """ border=""0"" alt=""" & _
                s_Caption & """ hspace=""10"" vspace=""2"" />"
            
        End If ' s_Image_Disabled = ""

    Else
    
        ' Build a querystring
        lsQueryString = NewQueryString(Array("Page"))
    
        ' Check to see if the built querystring has any data.
        If lsQueryString = "" Then
            
            ' No, add "page" field
            lsQueryString = "?Page="
            
        Else
        
            ' Yes, add querystring and "page" field
            lsQueryString = "?" & lsQueryString & "Page="
            
        End If ' lsQueryString = ""
    
        ' Create the <A> tag.
        lsResult = "<A href=""" & gsURL_ThisPage & lsQueryString & n_Page & _
            """ style=""color:" & sFGCOLOR_NAVIGATION_LINK & ";"">"
    
        ' Was an image specified?
        If s_Image = "" Then
            
            ' No, create a text link
            lsResult = lsResult & s_Caption
            
        Else
            
            ' yes, add the image.
            lsResult = lsResult & "<IMG src=""" & sIMAGE_ROOT & s_Image & """ " & _
                "border=""0"" alt=""" & s_Caption & """ hspace=""10""" & _
                " vspace=""2"" />"
            
        End If ' s_Image = ""
        lsResult = lsResult & "</A>"

    End If ' b_IsDisabled
    
    ' Pass results back
    HTML_NavLink = lsResult

End Function    
'———————————————————————————————————————————————————————————————————————————————
Private Function NewQueryString(ByRef UnwantedFields) ' As String

    Dim lsFieldName
    Dim lsBadName
    Dim lsResult
    Dim lbFoundBadName

    For Each lsFieldName In Request.QueryString
        lbFoundBadName = False
        For Each lsBadName In UnwantedFields
            If LCase(lsBadName) = LCase(lsFieldName) Then
                lbFoundBadName = True
                Exit For
            End If
        Next
        If Not lbFoundBadName Then
            lsResult = lsResult & lsFieldName & "=" & _
                Server.URLEncode(Request.QueryString(lsFieldName)) & "&"
        End If
    Next
    NewQueryString = lsResult
End Function
'———————————————————————————————————————————————————————————————————————————————
Sub Set_NextValue( _
    ByRef v_OriginalValue, _
    ByRef v_ValueListArray _
    )

    ' Sub Set_NextValue()
    '
    ' Routine reads original value and compares it to an array. After finding
    ' the current value in the array, it sets the variable to the next value
    ' in the array
    '
    ' Incoming Data
        ' Original value to search for
        ' List of values to alternate between
        '
        
    ' Declare local variables
    Dim lnIndex
    
    ' Loop through value list
    For lnIndex = LBound(v_ValueListArray, 1) To UBound(v_ValueListArray, 1)
    
        ' If the Original Value equals the current value in the array
        If v_OriginalValue = v_ValueListArray(lnIndex) Then
        
            ' Add 1 to the Index
            lnIndex = lnIndex + 1
            
            ' If the index is out of range
            If lnIndex > UBound(v_ValueListArray, 1) Then
                
                ' Set index to the first value in the array
                lnIndex = LBound(v_ValueListArray, 1)
            
            End If ' lnIndex > UBound(v_ValueListArray, 1)
            
            ' Set the Original Value to the index of the value list
            v_OriginalValue = v_ValueListArray(lnIndex)
            
            ' Exit the routine
            Exit Sub
        
        End If ' v_OriginalValue = v_ValueListArray(lnIndex)
    
    Next ' lnIndex
    
    ' Set the index to the first value in the array.
    lnIndex = LBound(v_ValueListArray, 1)
    
    ' Set the original value
    v_OriginalValue = v_ValueListArray(lnIndex)
    
End Sub ' Set_NextValue
'———————————————————————————————————————————————————————————————————————————————
Function SQL_Search( _
    ByVal s_Keywords, _
    ByRef s_FieldsSearchedArray, _
    ByRef s_MatchWords, _
    ByRef s_MatchType _
    )

    ' Function SQL_Select()
    '
    ' This routine outputs arguments for an SQL query depnding on
    ' information supplied by the calling routine.
    '
    ' Incomming Variables
        ' Keywords to look for.
        ' Array of fields to search.
        ' Weather keywords (shoud/should not) be found.
        ' Type of match (Any/All/Phrase)
        '
    
    ' Declare local variables
    Dim lnIndex                ' Represents an index (position) in an array

    Dim lsKeyWordArray        ' An array of keywords.
    Dim lsLogic                ' Logic to search for each keyword in database. (AND/OR)
    Dim lsResult            ' text that is returned form the procedure.
    

    ' If variable is not the correct variable type
    If Not IsArray(s_FieldsSearchedArray) Then
        
        ' Exit the procedure, returning nothing.
        Exit Function
    
    ' Else If no fields have been defined for searching
    ElseIf UBound(s_FieldsSearchedArray, 1) = -1 Then
    
        ' Exit the procedure, returning nothing.
        Exit Function

    ' Else If no keywords were defined
    ElseIf s_Keywords = "" Then
        
        ' Exit the procedure, returning nothing.
        Exit Function

    End If ' Not IsArray(s_FieldsSearchedArray)

    ' If query is to search records without the keywords
    If LCase(s_MatchWords) = "without" Then
        
        ' Add "NOT" to arguments
        lsResult = " NOT "

    End If ' s_MatchWords = "without"
    
    ' Start the SQL Query arguments as a seperate argument
    lsResult = lsResult & "("

    ' Double quote Aposterphies for SQL preperation.
    s_Keywords = Replace(s_Keywords, "'", "''")

    ' If keywords are to be matched as a phrase
    If LCase(s_MatchType) = "phrase" Then
    
        ' Add arguments to find phrase in at least 1 searchable field
        lsResult = lsResult & _
            Join(s_FieldsSearchedArray, " LIKE '%" & s_Keywords & "%' OR ") & _
            " LIKE '%" & s_Keywords & "%'"
    
    ' Else keywords are to be matched individually
    Else
        
        ' Convert all possible delimiters to commas
        s_Keywords = Replace(s_Keywords, " ", ",")    ' space
        s_Keywords = Replace(s_Keywords, ":", ",")    ' colen
        s_Keywords = Replace(s_Keywords, ";", ",")    ' semicolen
        s_Keywords = Replace(s_Keywords, "+", ",")    ' plus
        s_Keywords = Replace(s_Keywords, "-", ",")    ' minus
        s_Keywords = Replace(s_Keywords, """", ",")    ' Quote
        
        ' Loop while comma pairs exist
        While Not InStr(s_Keywords, ",,") = 0
        
            ' Replace comma pair with 1 comma
            s_Keywords = Replace(s_Keywords, ",,", ",")
            
        Wend ' Not InStr(s_Keywords, ",,") = 0
        
        ' Seperate keywords into an array with comma as delimeter
        lsKeyWordArray = Split(s_Keywords, ",")
                
        ' If all keywords must be matched
        If LCase(s_MatchType) = "all" Then
            
			If LCase(s_MatchWords) = "without" Then
				' Set SQL Logic to OR
				lsLogic = " OR "
			Else
				' Set SQL Logic to AND
				lsLogic = " AND "
			End If
        
        ' Else only 1 keyword needs to match
        Else
            
            ' Set SQL Logic to OR
            lsLogic = " OR "

        End If ' s_MatchType = "all"
        
        ' Loop through each keyword
        For lnIndex = 0 To UBound(lsKeyWordArray, 1)
            
            ' Add arguments to search keyword in each field
            lsResult = lsResult & Join(gsSearchFieldsArray, " LIKE '%" & _
                lsKeyWordArray(lnIndex) & "%' OR ") & " LIKE '%" & _
                lsKeyWordArray(lnIndex) & "%'"
                
            ' If we are not working with the last keyword
            If Not lnIndex = UBound(lsKeyWordArray, 1) Then
            
                ' Add logic to match (All/Any) of the keywords
                lsResult = lsResult & lsLogic

            End If ' Not lnIndex = UBound(lsKeyWordArray, 1)
            
        Next ' lnIndex

    End If ' s_MatchType = "phrase"
    
    ' Close argument
    lsResult = lsResult & ")"
    
    ' Return the results.
    SQL_Search = lsResult

End Function ' SQL_Search
'———————————————————————————————————————————————————————————————————————————————
Function SQL_Select( _
    ByRef b_IsSQL_DB, _
    ByRef s_FieldsReturnedArray, _
    ByRef s_Tables, _
    ByRef s_Keywords, _
    ByRef s_FieldsSearchedArray, _
    ByRef s_MatchWords, _
    ByRef s_MatchType, _
    ByRef s_Arguments, _
    ByRef s_SortField, _
    ByRef s_SortOrder _
    )
    
    ' Function SQL_Select()
    '
    ' This routine outputs an SQL SELECT query based on the 
    ' information supplied by the calling routine.
    '
    ' Incomming Variables
        ' Flag stating if query is for an SQL server
        ' Fields to request data from.
        ' Tables to request data from
        ' Keywords to search for
        ' Fields searched for keywords
        ' Weather keywords (shoud/should not) be found.
        ' Type of match (Any/All/Phrase)
        ' Arguments that the fields must pass to be returned.
        ' Name of field to be sorted
        ' Order in wich to sort the field
        '
    ' Dependent on routines
        ' SQL_Search()
        '
    ' --
    
    ' Declare local variables
    Dim lsResult            ' Results to output to calling routine.
    Dim lsSearchArguments    ' Arguments created by SQL_Search() function
    
    ' If querying an SQL Database
    If b_IsSQL_DB Then

        ' Tell the SQL server to select all records.
        lsResult = "SET ROWCOUNT 0 "

    End If ' End If b_SQL_DB

    ' Build the begining query.
    lsResult = lsResult & "SELECT "
    
    ' Add fields to retrieve
    lsResult = lsResult & Join(s_FieldsReturnedArray, ", ")
    
    ' Add the tables to retrieve fields from
    lsResult = lsResult & " FROM " & s_Tables
    
    ' Call SQL_Search() routine to return search arguments
    lsSearchArguments = SQL_Search( _
        s_Keywords, _
        s_FieldsSearchedArray, _
        s_MatchWords, _
        s_MatchType _
        )

    ' If search arguments were retrieved
    If Not lsSearchArguments = "" Then
        
        ' Add the search arguments
        lsResult = lsResult & " WHERE " & lsSearchArguments
    
    End If ' Not lsSearchArguments = ""

    ' If external arguments were received
    If Not s_Arguments = "" Then
    
        ' If WHERE clause has not been added yet
        If InStr(lsResult, s_Tables & " WHERE ") = 0 Then
    
            ' add WHERE clause
            lsResult = lsResult & " WHERE "
        
        ' Else add continuing statment
        Else
    
            ' add "AND" 
            lsResult = lsResult & " AND "
    
        End If ' End If InStr(lsResult, s_Tables & " WHERE ") = 0
    
        ' Add external arguments.
        lsResult = lsResult & "(" & s_Arguments & ")"
        
    End If ' End If Not s_Arguments = ""

    ' If a field is to be sorted
    If Not(s_SortField = "" Or s_SortOrder = "") Then

        ' Add " ORDER BY " clause
        lsResult = lsResult & " ORDER BY "
        
        ' Add field name and order
        lsResult = lsResult & s_SortField & " " & s_SortOrder

    End If 'End If Not(s_SortField = "" Or s_SortOrder = "")

    ' output results to calling routine.
    SQL_Select = lsResult
    
End Function
'———————————————————————————————————————————————————————————————————————————————
Const sDEBUG_PASSWORD = "abc123"

Sub WriteDebug( _
    ByRef s_Title, _
    ByRef s_Message _
    )

    ' Sub Write_Debug()
    '
    ' This routine receives a debuging message and checks to see if the clients
    ' cookie has been flagged to read them before. The output is a direct
    ' Write the the users web browser.
    '
    ' Incomming Variables
        ' Title, or type of debuging information
        ' Message to be shown
        '

    ' Define local variables
    Dim lsResult                ' Results to be created by routine
    
    Const lsHTMLFont = "<FONT face=""Helvetica,Arial,sans-serif"" size=""1"" color=""#000000"">"
    
    ' If the user does not have the ability to see debug information
    If Not Request.Cookies("Debug") = sDEBUG_PASSWORD Then
        
        ' Exit the routine
        Exit Sub
    
    End If ' Not Request.Cookies("Debug") = sDEBUG_PASSWORD
    
    ' Write Opening Comment
    lsResult = vbCrLf & "<!-- Debug " & s_Title & ": BEGIN -->" & vbCrLf
    
    ' Open Table and row
    lsResult = lsResult & "<TABLE class=""Debug"" bgcolor=""#cccccc"" " & _
        "width=""100%"">" & vbCrLf & "<TR>" & vbCrLf

    ' Write cell stating type of debug
    lsResult = lsResult & "<TD bgcolor=""#dddddd"" class=""cell-nolite"" " & _
        "align=""left"">" & vbCrLf & lsHTMLFont & vbCrLf & _
        "<B>Debugging Information - " & s_Title & "</B>" & vbCrLf & _
        "</FONT>" & vbCrLf & "</TD>" & vbCrLf
    
    ' Close row and open a new one.
    lsResult = lsResult & "</TR>" & vbCrLf & "<TR>" & vbCrLf
    
    ' Write cell with comments.
    lsResult = lsResult & "<TD bgcolor=""#eeeeee"" class=""cell-hilite"">" & _
        lsHTMLFont & vbCrLf & _
        Replace(Server.HTMLEncode(s_Message), vbCrLf, "<BR>") & vbCrLf & _
        "</FONT>" & vbCrLf & "</TD>" & vbCrLf
    
    ' Close the row and table.
    lsResult = lsResult & "</TR>" & vbCrLf & "</TABLE>" & vbCrLf & "<BR />" & _
        vbCrLf
    
    ' Write Closing Comment
    lsResult = lsResult & "<!-- Debug " & s_Title & ": END -->" & vbCrLf
    
    ' Write the resutls
    Response.Write(lsResult)

End Sub
'———————————————————————————————————————————————————————————————————————————————
' Dependant Routines:
    ' Sub InitializeGlobalVariables()
    
' Dependant Objects:
    ' Object(s) to request and retain data from a database.
        ' MDAC version 2.0 +
    '
Public Sub Main()

    Dim lsMatch1            ' records With/Without the keywords
    Dim lsMatch2            ' records containing Any/All/Phrase of Keywords
    
    Dim lbRendureCell        ' Are we going to rendure the current cell?

    Dim lnAbsolutePage        ' Current page Client is viewing
    Dim lnColumn            ' Colum Counter - Used for lvListDataArray
    Dim lnHighestRecordNum    ' Highest Record Number that will be viewed
    Dim lnIndex                ' Loop Counter
    Dim lnLinkColumnIndex    ' Location of column that is to be linked.
    Dim lnLowestRecordNum    ' Lowest Record Number that will be viewed
    Dim lnPageCount            ' Last page that contains data
    Dim lnRow                ' Row Counter - Used for lvListDataArray
    Dim lnTotalRecords        ' Total records that will be returned by database.

    Dim lsBgColor            ' Alternating Hex Background Color
    Dim lsClassName            ' Alternating StyleSheet Class Name
    Dim lsFgColor            ' Alternating Hex Foreground Color
    Dim lsFieldNameArray    ' List of real field names returned from query.
    Dim lsFieldName            ' Name of current field
    Dim lsHTML_Header        ' Table Header code to pass back to web browser.
    Dim lsHTML_ResultInfo    ' Some details about what was found (HTML)
    Dim lsHTML_Search        ' Search Form to find keywords (HTML)
    Dim lsLinkColor            ' Alternating Hex Link Color
    Dim lnPrimaryKeyIndex    ' The column that the primary key is located at.
    Dim lsQuery                ' Keywords that the client is searching for.
    Dim lsQueryString        ' Querystring to write to a link.
    Dim lsResult            ' Records returned from search (HTML)
    Dim lsSort                ' Sort Order to use on a field.
    Dim lsSortField            ' Field Name in database to sort.
    Dim lsSQL                ' SQL String to be executed by command object

    Dim lvListDataArray        ' All data contained within fields.

    ' Call routing to store user settings into global variables: Initailize Global Variables
    Call InitializeGlobalVariables()

    ' Store Default value of Link Column Index as Not found
    lnLinkColumnIndex = -1

    ' Store Default value of Primary Key Index as Not found
    lnPrimaryKeyIndex = -1
    
    ' Store Incomming Data (Page, Query, Sort, SortField)
    lnAbsolutePage    = Request.QueryString("Page")
    lsQuery            = Request.QueryString("Query")
    lsSort            = Request.QueryString("Sort")
    lsSortField        = Request.QueryString("SortField")
    lsMatch1        = Request.QueryString("Match1")
    lsMatch2        = Request.QueryString("Match2")

    ' If absolute page is not a number
    If lnAbsolutePage = "" Or Not IsNumeric(lnAbsolutePage) Then
    
        ' Store default absolute page value as 1
        lnAbsolutePage = 1
    
    ' Else absolute page was received correctly
    Else
    
        ' Convert absolute page to a number
        lnAbsolutePage = CInt(lnAbsolutePage)
    
    End If ' lnAbsolutePage = "" Or Not IsNumeric(lnAbsolutePage)

    ' If Absolute page is less then 1
    If lnAbsolutePage < 1 Then
        
        ' Change absolute page to 1
        lnAbsolutePage = 1
    
    End If ' lnAbsolutePage < 1

    ' ---------- RETRIEVE DATA
    
    ' Store SQL String from routine to find data: SQL_Select()
    lsSQL = SQL_Select( _
        bSQL_DATABASE, _
        gsDisplayFieldsArray, _
        gsRecordSource, _
        lsQuery, _
        gsSearchFieldsArray, _
        lsMatch1, _
        lsMatch2, _
        gsArguments, _
        lsSortField, _
        lsSort)

    ' Call on routine to connect to database and set variables: Data_GetListData()
        ' Pass SQL String
        ' Pass Records Per Page
        ' Pass Absolute Page to be viewed
        ' Sets TotalRecords returned
        ' Sets Page Count
        ' Sets Field Names returned from database
        ' Sets List data for the page
        ' Sets lowest record number displayed
        ' Sets highest record number displayed
    Call Data_GetListData( _
        lsSQL, _
        nRECORDS_PER_PAGE, _
        lnAbsolutePage, _
        lnTotalRecords, _
        lnPageCount, _
        lsFieldNameArray, _
        lvListDataArray, _
        lnLowestRecordNum, _
        lnHighestRecordNum _
        )

    ' Store an HTML Search Form generated by routine: HTML_Search()
        ' Pass Query
        ' Pass Match1
        ' Pass Match2
        ' Pass Fields to be retained
        ' Pass Image to be used as search button
        ' Returns HTML Search Form
    lsHTML_Search = HTML_Form_Search( _
        lsQuery, _
        lsMatch1, _
        lsMatch2, _
        gsRetainFieldsArray, _
        sIMAGE_BUTTON_SEARCH _
        )

    ' Store an HTML Result information generated by routine: HTML_ResultInfo()
        ' Pass Total Records Found
        ' Pass Absolute Page being viewed
        ' Pass Page Count
        ' Pass Lowest Record Number
        ' Pass Highest Record Number
        ' Returns HTML Result Information
    lsHTML_ResultInfo = HTML_BuildInformation( _
        lnTotalRecords, _
        lnAbsolutePage, _
        lnPageCount, _
        lnLowestRecordNum, _
        lnHighestRecordNum)
    
    ' ---------- DATA TABLE
    
    ' Start the data table code.
    lsResult = "<!-- Data Table: BEGIN -->" & vbCrLf

    ' Begin building the data table.

	lsResult = "<TABLE width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & _
		"<TR><TD bgcolor=""" & sBORDER_COLOR & """><TABLE width=""100%"" border=""0"" " & _
		"cellpadding=""0"" cellspacing=""1"" bgColor=""" & sBORDER_COLOR & """>" & vbCrLf 
    
    ' is a URL provided to delete records?
    If Not gsURL_DELETE = "" Then

        ' Write the <FORM id=form1 name=form1> to allow multiple records to be deleted.
        lsResult = lsResult & "<FORM method=""post"" action=""" & _
            gsURL_DELETE & "?" & NewQueryString(Array()) & """ name=""DeleteForm"">" & vbCrLf

    End If ' Not gsURL_DELETE = ""


    ' If data was returned for the current page in the database
    If IsArray(lvListDataArray) Then

        ' Build the header
        lsHTML_Header = HTML_BuildHeader( _
            lsFieldNameArray, _
            lnPrimaryKeyIndex, _
            lnLinkColumnIndex _
            )
    
        lsResult = lsResult & lsHTML_Header
    
        ' ---------- LIST-DATA: BEGIN
    
        ' Write opening List Data comments
        lsResult = lsResult & "<!-- List-Data: BEGIN -->" & vbCrLf
    
        ' Loop through rows on a 0 based array.
        For lnRow = 0 To UBound(lvListDataArray, 2)

            ' Set Alternate cell properties
            Call Set_NextValue(lsBgColor, sCellBgColorArray)
            Call Set_NextValue(lsFgColor, sCellFgColorArray)
            Call Set_NextValue(lsLinkColor, sCellFgColorLinkArray)
            Call Set_NextValue(lsClassName, sCellClassArray)

            ' Write opening Row coments
            lsResult = lsResult & vbCrLf & "<!-- Row " & (lnRow + 1) & ": BEGIN -->" & vbCrLf
            
            ' Begin a new row
            lsResult = lsResult & "<TR>" & vbCrLf
            
            ' Loop through each column/field.
            For lnColumn = 0 To UBound(lsFieldNameArray, 1)
            
                ' By default, we set each cell to be rendured
                lbRendureCell = True

                ' If the current column is the prmary key
                If lnColumn = lnPrimaryKeyIndex Then
                
                    ' If the linking column is not the primary key
                    If Not lnLinkColumnIndex = lnPrimaryKeyIndex Then
                        
                        ' Set Flag so that column will not be rendured
                        lbRendureCell = False

                    End If ' Not lnLinkColumnIndex = lnPrimaryKeyIndex

                End If ' lnColumn = lnPrimaryKeyIndex
            
                ' If the cell may be rendured
                If lbRendureCell Then
                
                    ' Add new cell to results
                    lsResult = lsResult & "<TD bgcolor=""" & lsBgColor & _
                        """ class=""" & lsClassName & """"
                    
                    ' Determine type of variable being looked at
                    Select Case VarType(lvListDataArray(lnColumn, lnRow))
                        
                        ' CASE (Number|Date|Currency)
                        Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, _
                            vbDate, vbCurrency
                        
                            ' Add cell property to align text to the right
                            lsResult = lsResult & " align=""right"""
                        
                        ' CASE (Boolean|Null)
                        Case vbBoolean, vbNull
                        
                            ' Add cell property to align text in the center
                            lsResult = lsResult & " align=""center"""

                        ' CASE (Text|Else)
                        Case Else
                        
                            ' Add cell property to align text to the left
                            lsResult = lsResult & " align=""left"""

                    End Select ' VarType(lvListDataArray(lnColumn, lnRow))

                    ' Add remining table cell tag
                    lsResult = lsResult & ">" & vbCrLf
                    
                    ' Add font style to results
                    lsResult = lsResult & "<FONT size=""" & sFONT_SIZE & """ " & "face=""" & _
                        sFONT_FACE & """ color=""" & lsFgColor & """>" & vbCrLf
                
                    ' If the column is to be linked, and we have a page to link to
                    If lnColumn = lnLinkColumnIndex And Not gsURL_Link = "" Then

                        ' Add a link
                        lsResult = lsResult & "<A href=""" & gsURL_Link
                        
                        ' If a querystring identifier is not present in the page to link to
                        If InStr(gsURL_Link, "?") = 0 Then
                            
                            ' Add Querystring Identifier
                            lsResult = lsResult & "?"
                        
                        ' Else a querystring identifier is present
                        Else
                        
                            ' Add Identifier for next field/value pair
                            lsResult = lsResult & "&"
                        
                        End If ' InStr(gsURL_Link, "?") = 0
                        
                        ' Add the primary key name/value pair to pass to the next record.
                        lsResult = lsResult & gsPrimaryKey & "=" & _
                            lvListDataArray(lnPrimaryKeyIndex, lnRow)
                        
                        ' Loop through fields that Webmaster wants to retain
                        For Each lsFieldName In gsRetainFieldsArray

                            ' Add additional field name/value pair
                            lsResult = lsResult & "&" & lsFieldName & "=" & _
                                Server.URLEncode(Request.QueryString(lsFieldName))

                        Next ' lsFieldName
                        
                        ' Add remainder of opening <A> tag to results
                        lsResult = lsResult & """ style=""color:" & lsLinkColor & ";"">"
                        
                    End If ' lnColumn = lnLinkColumnIndex And Not gsURL_Link = ""
                    
                    ' If the data is a NULL value
                    If IsNull(lvListDataArray(lnColumn, lnRow)) Then
                        
                        ' Add text signifying we have a null value
                        lsResult = lsResult & _
                            "<SPAN class=""List-Null"" style=""color:" & _
                            sFGCOLOR_NULL & ";"">&lt;NULL&gt;</SPAN>"
                    
                    ' ElseIf the data is EMPTY
                    ElseIf lvListDataArray(lnColumn, lnRow) = "" Then
                    
                        ' Add text signifying we have an Empty value
                        lsResult = lsResult & _
                            "<SPAN class=""List-Empty"" style=""color:" & _
                            sFGCOLOR_EMPTY & ";"">&lt;EMPTY&gt;</SPAN>"
                    
                    ' ElseIf the data is a DATE
                    ElseIf IsDate(lvListDataArray(lnColumn, lnRow)) Then
                    
                        ' Add data as a short date
                        lsResult = lsResult & _
                            FormatDateTime(lvListDataArray(lnColumn, lnRow), vbShortDate)

                    ' ElseIf the data is currency
                    ElseIf VarType(lvListDataArray(lnColumn, lnRow)) = vbCurrency Then
                        
                        ' Add data as dollar amount
                        lsResult = lsResult & FormatCurrency(lvListDataArray(lnColumn, lnRow), 2, True, False, True)
                    
                    
                    ElseIf VarType(lvListDataArray(lnColumn, lnRow)) = vbBoolean Then
                    
						If lvListDataArray(lnColumn, lnRow) Then
							lsResult = lsResult & "Yes"
						Else
							lsResult = lsResult & "No"
						End If

                    ElseIf IsNumeric(lvListDataArray(lnColumn, lnRow)) Then
						
						' Determine type of variable being looked at
						Select Case VarType(lvListDataArray(lnColumn, lnRow))
						    
						    ' CASE (Number|Date|Currency)
						    Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency

								If lvListDataArray(lnColumn, lnRow) < 0 Then
									lsResult = lsResult & "<FONT color=""#800000"">"
								End If
						
								' If the number is a decimal, limit it to 2
								If Int(lvListDataArray(lnColumn, lnRow)) = lvListDataArray(lnColumn, lnRow) Then
									lsResult = lsResult & FormatNumber(lvListDataArray(lnColumn, lnRow), 0, True, True, True)
								ElseIf Int(lvListDataArray(lnColumn, lnRow) * 10) = lvListDataArray(lnColumn, lnRow) * 10 Then
									lsResult = lsResult & FormatNumber(lvListDataArray(lnColumn, lnRow), 1, True, True, True)
								Else
									lsResult = lsResult & FormatNumber(lvListDataArray(lnColumn, lnRow), 2, True, True, True)
								End If

								If lvListDataArray(lnColumn, lnRow) < 0 Then
									lsResult = lsResult & "<FONT color=""#800000"">"
								End If
							Case Else
								lsResult = lsResult & lvListDataArray(lnColumn, lnRow)
						End Select
						
                    ' Else we have regular data
                    Else

                        ' Add data to results
                        lsResult = lsResult & lvListDataArray(lnColumn, lnRow)
                    
                    End If ' IsNull(lvListDataArray(lnColumn, lnRow))
                    
                    ' If the column is a linked column
                    If lnColumn = lnLinkColumnIndex Then
                    
                        ' Add closing <A> tag to results
                        lsResult = lsResult & "</A>"
                    End If

                    ' Add closing style and table cell tags to results
                    lsResult = lsResult & vbCrLf & _
                        "</FONT>" & vbCrLf & _
                        "</TD>" & vbCrLf

                End If ' lbRendureCell
                
            Next ' lnColumn


            ' If a URL has been defined for deleting records
            If Not gsURL_DELETE = "" Then
                
                ' Add opening table cell to results
                lsResult = lsResult & "<TD bgcolor=""" & lsBgColor & _
                    """ class=""" & lsClassName & """ align=""center"">" & _
                    vbCrLf
                
                ' Add font style properties to results
                lsResult = lsResult & "<FONT size=""2"" face=""" & _
                    sFONT_FACE & """>" & vbCrLf
                
                ' Add input(checkbox) for marking record to be deleted
                lsResult = lsResult & "<INPUT type=""checkbox"" name=""" & _
                    gsPrimaryKey & """ value=""" & _
                    lvListDataArray(lnPrimaryKeyIndex, lnRow) & """ />" & vbCrLf
                
                ' add closing fong-style and table cell tags to results
                lsResult = lsResult & "</FONT>" & vbCrLf & "</TD>" & vbCrLf
                
            End If ' Not gsURL_DELETE = ""
            
            ' Add closing row tag to results
            lsResult = lsResult & "</TR>" & vbCrLf

            ' Write closing Row coments to results
            lsResult = lsResult & "<!-- Row " & (lnRow + 1) & ": END -->" & vbCrLf

        Next ' lnRow

        ' ---------- LIST-DATA: END
    
        ' ---------- OPTIONS: BEGIN

        ' Write opening option coments
        lsResult = lsResult & "<!--Options: BEGIN -->" & vbCrLf

        ' Check to see if URL's are provided to add and delete records.
        If Not(gsURL_DELETE = "" And gsURL_ADD = "") Then
            
            ' Yes, begin a new row.
            lsResult = lsResult & "<TR>" & vbCrLf
            
            ' Open cell to add a button
            lsResult = lsResult & "<TD colspan="""
            
            ' Determine if the Primary Key is the column that is linked.
            If lnLinkColumnIndex = lnPrimaryKeyIndex Then
            
                ' Yes, Set the colspan = count of all field columns
                lsResult = lsResult & (UBound(gsDisplayFieldsArray, 1) + 1)

            Else

                ' No, Set the colspan = count of all field columns - 1
                lsResult = lsResult & (UBound(gsDisplayFieldsArray, 1))
            
            End If
            
            ' Finnish writing the opening cell
            lsResult = lsResult & """ align=""left"" bgcolor" & _
                "=""" & sBGCOLOR_CELL_HEADER & """ class=""Cell-Header"">" & vbCrLf
            
            ' Open the style of text.
            lsResult = lsResult & "<FONT color=""" & sFGCOLOR_CELL_HEADER & _
                """ face=""" & sFONT_FACE & """>" & vbCrLf
            
            ' Determine if a URL was provided for adding records.
            If gsURL_ADD = "" Then
            
                ' No, write an invisible small character.
                lsResult = lsResult & "<FONT color=""" & sBGCOLOR_CELL_HEADER & _
                    """ size=""1"" style=""NoAdd"">x</FONT>" & vbCrLf
            
            Else
            
                ' Yes, Write link to page to add records.
                lsResult = lsResult & "<A href=""" & gsURL_ADD & """ " & _
                    "style=""color:" & sFGCOLOR_CELL_HEADER_LINK & ";"">"
            
                ' Determine if a button was defined for adding records.    
                If sIMAGE_BUTTON_ADD = "" Then
                
                    ' No, add a text link to add a record.
                    lsResult = lsResult & "Add Record"

                Else
            
                    ' Add an image linking to page to add record.
                    lsResult = lsResult & "<IMG src=""" & _
                        sIMAGE_ROOT & sIMAGE_BUTTON_ADD & """ border=""0"" " & "vspace=""2""" & _
                        " hspace=""10"" alt=""Add Record"" />"
                
                End If ' sIMAGE_BUTTON_ADD = ""
                
                ' Close the link
                lsResult = lsResult & "</A>" & vbCrLf

            End If ' gsURL_ADD = ""

            ' Close the table cell for "add record" link.
            lsResult = lsResult & "</FONT>" & vbCrLf & "</TD>" & vbCrLf
            
            
            ' Open the delete cell
            lsResult = lsResult & "<TD align=""center"" bgcolor=""" & _
                sBGCOLOR_CELL_HEADER & """ class=""Cell-Header"">" & vbCrLf
            
            ' Determine if a URL was provided for adding records.
            If gsURL_DELETE = "" Then
                
                ' No, write an invisible small character.
                lsResult = lsResult & "<FONT color=""" & sBGCOLOR_CELL_HEADER & _
                    """ size=""1"" style=""NoAdd"">x</FONT>" & vbCrLf

            Else
            
                ' Determine if a button was defined for deleting records.
                If sIMAGE_BUTTON_DELETE = "" Then
                
                    ' No, write a standard Submit input button
                    lsResult = lsResult & "<INPUT type=""submit"" " & _
                        "value=""Delete"" />" & vbCrLf
            
                Else

                    ' Yes, Write image input button.
                    lsResult = lsResult & "<INPUT type=""image"" " & _
                        "value=""Delete"" src=""" & sIMAGE_ROOT & sIMAGE_BUTTON_DELETE & """" & _
                        " border=""0"" vspace=""3"" hspace=""3"" " & _
                        "alt=""Delete Marked Records"" />" & vbCrLf
                
                End If ' sIMAGE_BUTTON_DELETE = ""
            
            End If ' gsURL_DELETE = ""
            
            ' Close Cell, Style, and form tags.
            lsResult = lsResult & _
                "</TD>" & vbCrLf & _
                "</TR>" & vbCrLf & _
                "</FORM>" & vbCrLf
                
        End If ' Not(gsURL_DELETE = "" And gsURL_ADD = "")
    
        ' Write closing option comments
        lsResult = lsResult & "<!--Options: END -->" & vbCrLf
    
        ' ---------- OPTIONS: END

        ' Close table that holds column headers, list data, and options.
        lsResult = lsResult & "</TABLE></TD></TR></TABLE>" & vbCrLf

	Else
		lsResult = "<FONT face=""" & sFONT_FACE & """ size=""" & sFONT_SIZE & """>No Data Available</FONT><BR>"
    End If ' IsArray(lvListDataArray)

    ' Write out the results.
    If gbShowSearch Then Response.Write(lsHTML_Search)

    If gbShowResultInfo Then Response.Write(lsHTML_ResultInfo)
    
    Response.Write(lsResult)

    Response.Write(HTML_BuildNavigation(lnAbsolutePage, lnPageCount))
    
End Sub ' WriteList
'———————————————————————————————————————————————————————————————————————————————
%>
