<%

' Common Debug routines to see incomming data

Dim sKey
Dim nItemCount
Dim nItemIndex

If Request.QueryString = "" Then
	Response.Write "<BR>QueryString data is not present"
Else
	Response.Write "<BR>QueryString data:"
	
	Response.Write("<OL>")
	For Each sKey In Request.QueryString
		Response.Write("<LI>")
		Response.Write(sKey & " = ")
		nItemCount = Request.QueryString(sKey).Count
		If nItemCount = 1 Then
			Response.Write("""" & Request.QueryString(sKey) & """")
		Else
			Response.Write("<OL>")
			For nItemIndex = 1 To nItemCount
				Response.Write("<LI>""" & Request.QueryString(sKey)(nItemIndex) & """</LI>")
			Next
			Response.Write("</OL>")
		End If
		Response.Write "</LI>"
	Next
	Response.Write("</OL>")
End If


If Request.Form = "" Then
	Response.Write "<BR>Form data is not present"
Else
	Response.Write "<BR>Form data:"
	
	Response.Write("<OL>")
	For Each sKey In Request.Form
		Response.Write("<LI>")
		Response.Write(sKey & " = ")
		nItemCount = Request.Form(sKey).Count
		If nItemCount = 1 Then
			Response.Write("""" & Request.Form(sKey) & """")
		Else
			Response.Write("<OL>")
			For nItemIndex = 1 To nItemCount
				Response.Write("<LI>""" & Request.Form(sKey)(nItemIndex) & """</LI>")
			Next
			Response.Write("</OL>")
		End If
		Response.Write "</LI>"
	Next
	Response.Write("</OL>")
End If
%>