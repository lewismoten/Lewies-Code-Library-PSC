<%Option Explicit%>
<!--#INCLUDE FILE="Common.inc"-->
<%
Dim ObjSite
Call Authenticate(ObjSite)
%>
<HTML>
	<HEAD>
		<TITLE>Custom Error Mapping</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Default.css">
		<SCRIPT Language="JavaScript" type="text/javascript" src="Default.js"></SCRIPT>
	</HEAD>
	<BODY>
		<SCRIPT>
		document.write(HeaderGUI('<%=Request.ServerVariables("SERVER_NAME")%>', '<%=Request.ServerVariables("APPL_MD_PATH")%>', '<%=ObjSite.KeyType%>'));
		<%=GetErrorList(ObjSite)%>
		document.write(FooterGUI());
		</SCRIPT>
	</BODY>
</HTML>

<%
Set ObjSite = Nothing
' ------------------------------------------------------------------------------
Function GetErrorList(ByRef pObjSite)

	Dim lObjInfo
	Dim lStrErrorsAry
	Dim lStrErrorAry
	Dim lStrMappingAry
	Dim lLngErrorIndex
	Dim lLngMaxErrorIndex
	
	Set lObjInfo = GetObject("IIS://localhost/w3svc/info")
	lStrErrorsAry = lObjInfo.GetEx("CustomErrorDescriptions")
	Set lObjInfo = Nothing

	lStrMappingAry = pObjSite.GetEx("HTTPErrors")

	lLngMaxErrorIndex = UBound(lStrErrorsAry)

	For lLngErrorIndex = 0 To lLngMaxErrorIndex
		lStrErrorAry = Split(lStrErrorsAry(lLngErrorIndex), ",")
		GetErrorList = GetErrorList & GetErrorRecord(lStrErrorAry, lStrMappingAry)
	Next

End Function
' ------------------------------------------------------------------------------
Function GetErrorRecord(ByRef pStrErrorAry, ByRef pStrMappingAry)

	Dim lStrErrorCode
	Dim lStrType
	Dim lStrDefaultText
	Dim lStrContents
	Dim lLngMapIndex
	Dim lStrMapAry
	Dim lStrPath
	
	
	lLngMapIndex = FindMapIndex(pStrErrorAry, pStrMappingAry)
	
	If lLngMapIndex = -1 Then
		lStrMapAry = Array("","","","")
	Else
		lStrMapAry = Split(pStrMappingAry(lLngMapIndex), ",")
	End If

	GetErrorRecord = GetErrorRecord & _
		"document.write(ErrorGUI(" & _
			pStrErrorAry(0) & ", " & _
			pStrErrorAry(1) & ", " & _
			"'" & pStrErrorAry(2) & "', " & _
			"'" & pStrErrorAry(3) & "', " & _
			"'" & lStrMapAry(2) & "', " & _
			"'" & lStrMapAry(3) & "'));" & vbCrLf
			
End Function
' ------------------------------------------------------------------------------
Function FindMapIndex(ByRef pStrErrorAry, ByRef pStrMappingAry)
	Dim lLngMaxIndex
	Dim lLngIndex
	Dim lStrMapAry
	lLngMaxIndex = UBound(pStrMappingAry)
	For lLngIndex = 0 To lLngMaxIndex
		lStrMapAry = Split(pStrMappingAry(lLngIndex), ",")
		If lStrMapAry(0) = pStrErrorAry(0) Then
			If lStrMapAry(1) = pStrErrorAry(1) Then	
				FindMapIndex = lLngIndex
				Exit Function
			End If
		End If
	Next
	FindMapIndex = -1
End Function
' ------------------------------------------------------------------------------
%>