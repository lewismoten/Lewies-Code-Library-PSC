<%Option Explicit%>
<!--#INCLUDE FILE="Inc\clsDatabase.asp"-->
<!--#INCLUDE FILE="Inc\clsDatalist.asp"-->
<LINK rel="stylesheet" type="text/css" href="style/datalist.css">
<%

Dim oDatabase
Dim oDatalist
Dim sTable

Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist


oDataList.urlAdd = "Countries.Add.asp"
oDataList.urlEdit = "Countries.Edit.asp"
oDataList.urlDelete = "Countries.Delete.asp"
oDataList.urlInsert = "Countries.Insert.asp"
oDataList.urlMoveUp = "Countries.MoveUp.asp"
oDataList.urlMoveDown = "Countries.MoveDown.asp"

oDataList.PageSize = 5	' Default is 10

oDataList.PrimaryKey = "CountryID"
oDataList.TableName = "Countries"
oDataList.DefaultSortedField = "CountryName"
oDataList.DefaultSortedOrder = "ASC"

oDataList.AddSearchedField "CountryName"
oDataList.AddSearchedField "CountryCode2"
oDataList.AddSearchedField "CountryCode3"

oDataList.AddReturnedField "CountryName", "Name"
oDataList.AddReturnedField "CountryCode3", "Code (3)"
oDataList.AddReturnedField "CountryCode2", "Code (2)"

sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
%>