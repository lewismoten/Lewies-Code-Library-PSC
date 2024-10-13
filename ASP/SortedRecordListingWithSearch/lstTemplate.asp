<%
Option Explicit
Response.Buffer = True
Response.Expires = 0
%>
<!--#INCLUDE VIRTUAL="/inc/database.asp"-->
<!--#INCLUDE VIRTUAL="/inc/lstRecords.asp"-->
<meta HTTP-EQUIV="Expires" CONTENT="Thu, 01 Dec 1994 16:00:00 GMT">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<%

' -----
' On Error Resume Next

Call Main()

'———————————————————————————————————————————————————————————————————————————————
Sub InitializeGlobalVariables()
    ' ------------------------------
    ' Basic Query Settings
    ' -----

    ' Table(s) to search from
        ' "Users"
        ' "Books INNER JOIN Authors ON Books.BookID = Authors.AuthorID"
	gsRecordSource = "COM_Communities"

    ' Set Primary Key
    gsPrimaryKey = "CommunityID"

    ' Set arguments to look for in SQL String 
        ' "SELECT * FROM TABLE WHERE " & gcArguments
        ' Null defaults to no arguments.
        ' "" defaults to no arguments.
    gsArguments = Null

    ' Set the fields to be displayed.
        ' gsDisplayFieldsArray = Array("BookID", "Title", "Author")
    gsDisplayFieldsArray = Array( _
        "CommunityID", _
        "CommunityName", _
        "FileName" _
        )

    ' Set the captions to be displayed for each column
        ' IMPORTANT! IMPORTANT! IMPORTANT! IMPORTANT! IMPORTANT! IMPORTANT!
                ' MATCHES gsDisplayFieldsArray FIELDS.
                ' MAKE SURE SUBSCRIPT IS WITHIN RANGE!
        ' gsCaptionsAry = Array("BookID", "Title", "Author")
        
    gsCaptionsAry = Array( _
        "ID", _
        "Community", _
        "Folder Name" _
        )
    ' ------------------------------
    ' User Query Settings
    ' -----

    ' Set the fields to be searched.
        ' Array("Title", "Author")
        ' Array() ' To disable feature
    gsSearchFieldsArray = Array( _
        "CommunityName", _
        "FileName" _
        )

    ' Set the fields to pass along from the querystring.
        ' Array("AuthorID")        ' To request and pass AuthorID.
        ' Array()                ' To disable feature
    gsRetainFieldsArray = Array()

    ' ------------------------------
    ' Display Settings
    ' -----

    ' set field name to link to other pages
        ' Null defaults to gsPrimaryKey
        ' "" defaults to gsPrimaryKey
        ' * If Not(gsLinkField = gsPrimaryKey) Then gsPrimaryKey removed from view
    gsLinkField = "CommunityName"

    ' Determine if search form is displayed.
    gbShowSearch = True

    ' Determine if result information is displayed.
    gbShowResultInfo = True

    ' ------------------------------
    ' Navigation 
    ' -----

    ' URL - this page
    gsURL_ThisPage = "lstCommunities.asp"

    ' To disable any of the following URLs, set the value to Null

    ' URL - link records (to view or edit)
    gsURL_Link = "edtCommunity.asp"

    ' URL - add record
    gsURL_Add = "addCommunity.asp"

    ' URL - delete record
    gsURL_Delete = "delCommunities.asp"

End Sub
'———————————————————————————————————————————————————————————————————————————————
%>
