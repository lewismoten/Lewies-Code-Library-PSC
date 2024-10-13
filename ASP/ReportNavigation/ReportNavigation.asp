<%
    '**************************************
    ' Name: Report Navigation
    ' Description:Gives you an interface to 
    '     navigate reports by year, quarter, month
    '     , week, or day. This is just the navigat
    '     ion alone that I append to the top of di
    '     fferent reports.
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
    
    Sub WriteReportNav(ByRef pstrView, ByRef pdtmDate)
    	Dim lstrOptions
    	Dim ldtmBegin
    	Dim ldtmEnd
    	
    	If dtmDate = "" Then dtmDate = Request.Cookies("Date")
    	If strView = "" Then strView = Request.Cookies("View")
    	If pdtmDate = "" Or Not IsDate(pdtmDate) Then pdtmDate = Date()
    	Select Case pstrView
    		Case "y", "ww", "m", "q", "yyyy"
    		Case Else
    			pstrView = "m"
    	End Select
    	Response.Cookies("View") = strView
    	Response.Cookies("Date") = dtmDate
    	lstrOptions = "y,Day;ww,Week;m,Month;q,Quarter;yyyy,Year"
    	lstrOptions = Replace(lstrOptions, ",", """>")
    	lstrOptions = Replace(lstrOptions, ";", "</OPTION><OPTION value=""")
    	lstrOptions = "<OPTION value=""" & lstrOptions & "</OPTION>"
    	lstrOptions = Replace(lstrOptions, "value=""" & pstrView & """", "value=""" & pstrView & """ selected")
    	Response.Write "<TABLE width=""100%"">"
    	Response.Write "<TR>"
    	Response.Write "<TD align=""center"">"
    	Response.Write "<A href=""Default.asp?View=" & pstrView & "&Date=" & Server.URLEncode(DateAdd(pstrView, -1, pdtmDate)) & """>"
    	Response.Write "&lt;&lt; Back</A>"
    	Response.Write "</TD>"
    	Response.Write "<TD align=""center""><B>"
    	Select Case strView
    		Case "y"
    			Response.Write "Day " & DatePart("y", pdtmDate) & ": " & FormatDateTime(pdtmDate, vbLongDate)
    		Case "ww"
    			Response.Write "Week " & DatePart("ww", pdtmDate) & " of " & Year(pdtmDate) & "<BR>"
    			ldtmBegin = DateAdd("w", - (DatePart("w", pdtmDate) - 1), pdtmDate)
    			ldtmEnd = DateAdd("w", 7 - (DatePart("w", pdtmDate)), pdtmDate)
    			Response.Write MonthName(Month(ldtmBegin)) & " " & Day(ldtmBegin) & " through " & _
    			MonthName(Month(ldtmEnd)) & " " & Day(ldtmEnd)
    		Case "m"
    			Response.Write MonthName(DatePart("m", pdtmDate)) & " of " & DatePart("yyyy", pdtmDate)
    		Case "q"
    			Select Case DatePart("q", pdtmDate)
    				Case 1
    					Response.Write "1st Quarter of " & DatePart("yyyy", pdtmDate) & "<BR>January through March"
    				Case 2
    					Response.Write "2nd Quarter of " & DatePart("yyyy", pdtmDate) & "<BR>April through June"
    				Case 3
    					Response.Write "3rd Quarter of " & DatePart("yyyy", pdtmDate) & "<BR>July through September"
    				Case 4
    					Response.Write "4th Quarter of " & DatePart("yyyy", pdtmDate) & "<BR>October through December"
    			End Select
    		Case "yyyy"
    			Response.Write "Year of " & DatePart("yyyy", pdtmDate)
    	End Select
    	Response.Write "</B></TD>"
    	Response.Write "<TD align=""center"">"
    	Response.Write "<A href=""Default.asp?View=" & pstrView & "&Date=" & Server.URLEncode(DateAdd(pstrView, 1, pdtmDate)) & """>Next &gt;&gt;</A>"
    	Response.Write "</TD>"
    	Response.Write "<TD align=""right"">"
    	Response.Write "View: "
    	Response.Write "<SELECT name=""View"" onChange=""window.location.search='View=' + this[this.selectedIndex].value + '&Date=" & Server.URLEncode(pdtmDate) & "'"">"
    	Response.Write lstrOptions
    	Response.Write "</SELECT>"
    	Response.Write "</TD>"
    	Response.Write "</TR>"
    	Response.Write "</TABLE>"
    End Sub
    %>