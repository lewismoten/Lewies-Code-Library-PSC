<!--#INCLUDE FILE="Common.inc"-->
<%
' Constants
Const ADS_PROPERTY_CLEAR       = 1
Const ADS_PROPERTY_UPDATE      = 2
Const ADS_PROPERTY_APPEND      = 3
Const ADS_PROPERTY_DELETE      = 4

Dim lLngErrorCode
Dim lStrCustomType
Dim lStrCustomPath
Dim lStrErrorAry

Dim ObjSite

' Make sure someone with proper permissions can update information
Call Authenticate(ObjSite)

' If nothing was received - redirect to default page
If Request.Form = "" Then Response.Redirect("default.asp")

LngCount = Request.Form("ErrorCode").Count

' Loop through each error code
For LngItem = 1 To LngCount

	' Grab posted form data
	lLngErrorCode = Request.Form("ErrorCode")(LngItem)
	lStrCustomType = Request.Form("CustomType")(LngItem)
	lStrCustomPath = Request.Form("CustomPath")(LngItem)
	If Not (lStrCustomType = "" Or lStrCustomPath = "") Then 

		' Append the set of data delimited by commas
		lStrError = lStrError & _
			lLngErrorCode & "," & _
			lStrCustomType & "," & _
			lStrCustomPath

		' append a pipe to deliminate between data sets
		If Not LngItem = LngCount Then
			lStrError = lStrError & "|"
		End If

	End If
Next

' split data sets by the pipe into an array.
lStrErrorAry = Split(lStrError, "|")

' Update errors
Call ObjSite.PutEx(ADS_PROPERTY_UPDATE, "HTTPErrors", lStrErrorAry)

' Commit the changes in the cache to the underlying
' directory store to make the changes persistent. 
Call ObjSite.SetInfo()

Set ObjSite = Nothing ' Release Object

' Go back to the list
Response.Redirect("default.asp")
%>