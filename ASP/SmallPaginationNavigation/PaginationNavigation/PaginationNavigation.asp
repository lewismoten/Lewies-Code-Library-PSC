<%
Function PaginationNavigation(ByVal pLngPage, ByRef pLngPageCount, ByVal pStrURL, ByRef pLngMax)
	
	Dim lStrNav
	Dim lLngFirst
	Dim lLngLast
	Dim lLngIndex
	
	If pLngPageCount < 2 Then Exit Function
	If Not IsNumeric(pLngPage) Or pLngPage = "" Then pLngPage = 1
	pLngPage = CLng(pLngPage)
	If pLngPage < 1 Then pLngPage = 1
	
	lLngFirst = pLngPage - (pLngMax \ 2)
	lLngLast = lLngFirst + (pLngMax - 1)
	If lLngFirst < 1 Then
		lLngFirst = 1
		lLngLast = pLngMax
	End If
	
	If lLngLast > pLngPageCount Then
		lLngLast = pLngPageCount
		lLngFirst = lLngLast - (pLngMax - 1)
	End If
	
	If lLngFirst < 1 Then lLngFirst = 1
	
	lStrNav = "Page: "
	
	If InStr(1, pStrURL, "?") = 0 Then
		pStrURL = pStrURL & "?Page="
	Else
		If Right(pStrURL, 1) = "?" Then
			pStrURL = pStrURL & "Page="
		Else
			pStrURL = pStrURL & "&Page="
		End If
	End If
	
	If Not lLngFirst = 1 Then
		lStrNav = lStrNav & "<A href=""" & pStrURL & "1"">1</A> ... "
	End If
	
	For lLngIndex = lLngFirst To lLngLast
		If lLngIndex = pLngPage Then lStrNav = lStrNav & "["
		lStrNav = lStrNav & "<A href=""" & pStrURL & lLngIndex & """>" & lLngIndex & "</A>"
		If lLngIndex = pLngPage Then lStrNav = lStrNav & "]"
		lStrNav = lStrNav & " "
	Next
	
	If Not lLngLast = pLngPageCount Then
		lStrNav = lStrNav & " ... <A href=""" & pStrURL & pLngPageCount & """>" & pLngPageCount & "</A> "
	End If
	
	PaginationNavigation = lStrNav
	
End Function
%>