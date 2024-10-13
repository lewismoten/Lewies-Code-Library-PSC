<%
    '**************************************
    ' for :Input Date
    '**************************************
    'Copyright (C) 1999 by Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Input Date
    ' Description:This small snippet will ge
    '     nerate 3 drop down boxes that will allow
    '     your users to choose a date.
    ' By: Lewis Moten
    '
    '
    ' Inputs:asName - String representing th
    '     e name that you want the field to be pas
    '     sed as. The three fields will be passed 
    '     as [anName].Month, [anName].Day, [anName
    '     ].Year.
    'adDate - The current date that will be selected. This is parsed into the Month, Day, and Year if it is a valid date.
    '
    ' Returns:This function returns HTML cod
    '     e for three "SELECT" boxes next to each 
    '     other.
    '
    'Assumes:It is expected that you have op
    '     ened a FORM element before writing out t
    '     he results from this function. Also, you
    '     should know how to grab form data. An ad
    '     ditional function has been provided to p
    '     arse posted data from this function.
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    ' --------------------------------------
    '     ---------------------------------------
    ' Name:		Input Date Select Boxes
    ' Description:
    '		Allows users to choose a date from 3 
    '     <SELECT> fields
    '		(Month, Day, and Year)
    '
    ' Author:	Lewis Moten
    ' Email:	lewis@moten.com
    ' URL:		http://www.lewismoten.com
    ' --------------------------------------
    '     ---------------------------------------
    ' Returns 3 HTML Selection boxes to pick
    '     a date
    Function InputDate(ByRef asName, ByRef adDate)
    	Dim lnDay			' Day
    	Dim lnMonth			' Month
    	Dim lnYear			' Year
    	
    	Dim lnIndex			' Counter
    	
    	Dim lsHTML			' HTML to be returned
    	' If a valid date was received	
    	If IsDate(adDate) Then
    		' Parse the Month, Day, and Year from the date
    		lnMonth = Month(adDate)
    		lnDay = Day(adDate)
    		lnYear = Year(adDate)
    	Else
    		' Initialize Month, Day, and Year as not being chosen
    		lnMonth = -1
    		lnDay = -1
    		lnYear = -1
    	End If
    	' ----- MONTH
    	
    	' Create the month selection field
    	lsHTML = "<SELECT name=""" & asName & ".Month"">" & vbCrLf
    	
    	' Add empty option - say field is month
    	lsHTML = lsHTML & "<OPTION value="""">Month</OPTION>" & vbCrLf
    	
    	' Loop through each month
    	For lnIndex = 1 To 12
    		
    		' Create an option with the numeric value of the month
    		lsHTML = lsHTML & "<OPTION value=""" & lnIndex & """"
    		
    		' Select the option if it matches the parsed month
    		If lnIndex = lnMonth Then lsHTML = lsHTML & " selected"
    		
    		' Add the name of the month as the text of the option
    		lsHTML = lsHTML & ">" & MonthName(lnIndex, False) & "</OPTION>" & vbCrLf
    	
    	Next
    	
    	' Close the Month selection field
    	lsHTML = lsHTML & "</SELECT>" & vbCrLf
    	
    	' ----- DAY
    	
    	' Create the day selection field
    	lsHTML = lsHTML & "<SELECT name=""" & asName & ".Day"">" & vbCrLf
    	
    	' Add empty option - say field is day
    	lsHTML = lsHTML & "<OPTION value="""">Day</OPTION>" & vbCrLf
    	
    	' Loop through 1 to max days in a month
    	For lnIndex = 1 To 31
    		
    		' create option with numeric value of day
    		lsHTML = lsHTML & "<OPTION value=""" & lnIndex & """"
    		
    		' select option if it matches parsed day
    		If lnIndex = lnDay Then lsHTML = lsHTML & " selected"
    		
    		' add the value of the day as the text to the option
    		lsHTML = lsHTML & ">" & lnIndex & "</OPTION>" & vbCrLf
    		
    	Next
    	
    	' close the day selection field
    	lsHTML = lsHTML & "</SELECT>" & vbCrLf
    	' ----- YEAR
    	
    	' Create the year selection field
    	lsHTML = lsHTML & "<SELECT name=""" & asName & ".Year"">" & vbCrLf
    	
    	' Add empty option - say field is year
    	lsHTML = lsHTML & "<OPTION value="""">Year</OPTION>" & vbCrLf
    	
    	' Loop from highest 4 digit year to lowest.
    	' range is 5 years from now to 100 years ago
    	For lnIndex = (Year(Date()) + 5) To (Year(Date()) - 100) Step -1
    	
    		' create option with year
    		lsHTML = lsHTML & "<OPTION value=""" & lnIndex & """"
    		
    		' select option if year matches parsed year
    		If lnIndex = lnYear Then lsHTML = lsHTML & " selected"
    		
    		' write the year as the text of the option
    		lsHTML = lsHTML & ">" & lnIndex & "</OPTION>" & vbCrLf
    	Next
    	
    	' close the year selection field
    	lsHTML = lsHTML & "</SELECT>" & vbCrLf
    	
    	' -----
    	
    	' Return the HTML that creates the Date Picker
    	InputDate = lsHTML
    	
    End Function
    ' --------------------------------------
    '     ---------------------------------------
    ' Parse a date comming from InputDate co
    '     mbination fields
    Function GetInputDate(ByRef asName)
    	
    	Dim ldDate
    	Dim lnDay
    	Dim lnMonth
    	Dim lnYear
    	
    	lnDay = Request.Form(asName & ".Day")
    	lnMonth = Request.Form(asName & ".Month")
    	lnYear = Request.Form(asName & ".Year")
    	
    	ldDate = lnMonth & "/" & lnDay & "/" & lnYear
    	
    	If IsDate(ldDate) Then
    		GetInputDate = CDate(ldDate)
    	End If
    	
    End Function
    ' --------------------------------------
    '     ---------------------------------------
    %>