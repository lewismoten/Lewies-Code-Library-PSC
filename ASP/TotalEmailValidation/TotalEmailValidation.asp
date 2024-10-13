<%
    '**************************************
    ' for :Total Email Validation
    '**************************************
    'Copyright (c) 1999, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Total Email Validation
    ' Description:Validates email addresses.
    '     Makes sure the email addresses with IP a
    '     ddresses are not private network address
    '     es. Allows multiple sub-domain levels. v
    '     erifies characters within domain names. 
    '     only allows standard length 26 character
    '     s for each domain name level, except the
    '     top (3 max)
    ' By: Lewis Moten
    '
    '
    ' Inputs:asString - Email address to be 
    '     validated.
    '
    ' Returns:Boolean (true/false) value ind
    '     icating if the string presented was a va
    '     lid email address.
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    Function IsEmail(ByRef asString)
    	Dim lsDomain
    	Dim lsSubDomain
    	Dim lsSubDomainArray
    	Dim lbIsIPdomain
    	Dim lnStart
    	Dim lsUserName
    	Dim lnOctect
    	Dim lnOctect2
    	Dim lnIndex
    	Const lsDOMAIN_CHARACTERS = ".ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-"
    	' Must have at least 6 characters "a@a.ru"
    	If Len(asString) < 6 Then
    		IsEmail = False
    		Exit Function
    	End If
    	
    	' Look for "@" delimiter
    	If Not InStr(asString, "@") > 1 Then
    		IsEmail = False
    		Exit Function
    	End If
    	
    	' Make sure characters exist after the "@"
    	If Len(asString) = InStr(asString, "@") Then
    		IsEmail = False
    		Exit Function
    	End If
    	
    	' Grab domain information "a.ru"
    	lsDomain = UCase(Mid(asString, InStr(asString, "@") + 1))
    	
    	' Grab username information
    	lsUserName = UCase(Left(asString, InStr(asString, "@") - 1))
    	
    	' Make sure at least 1 "." exists
    	If InStr(lsDomain, ".") = 0 Then
    		IsEmail = False
    		Exit Function
    	End If
    	
    	' Check for valid domain characters
    	lnStart = 1
    	Do While lnStart <= Len(lsDomain)
    		If InStr(lsDOMAIN_CHARACTERS, Mid(lsDomain, lnStart, 1)) Then
    			lnStart = lnStart + 1
    		Else
    			IsEmail = False
    			Exit Function
    		End If
    	Loop
    	
    	' Split domains
    	lsSubDomainArray = Split(lsDomain, ".")
    	lbIsIPdomain = False
    		
    	' Loop through each domain
    	For lnIndex = 0 To UBound(lsSubDomainArray, 1)
    		lsSubDomain = lsSubDomainArray(lnIndex)
    		If Len(lsSubDomain) = 0 Then
    			IsEmail = False
    			Exit Function
    		End If
    		
    		' Check to see if the domain is an IP Address
    		If lnIndex = 0 Then
    			If IsNumeric(lsSubDomain) Then
    				
    				' Only IP Addresses can have only numbers in subdomain area
    				lbIsIPDomain = True
    				
    				' Make sure 4 subdomains are present
    				If Not UBound(lsSubDomainArray, 1) = 3 Then
    					IsEmail = False
    					Exit Function
    				End If
    			
    			End If
    		End If
    		
    		If lbIsIPDomain Then
    			If Len(lsSubDomain) > 3 Then
    				IsEmail = False
    				Exit Function
    			ElseIf Not InStr(lsSubDomain, "-") = 0 Then
    				IsEmail = False
    				Exit Function
    			ElseIf Not IsNumeric(lsSubDomain)Then
    				IsEmail = False
    				Exit Function
    			End If
    			
    			lnOctect = CInt(lsSubDomain)
    			
    			If lnOctect > 255 Then
    				IsEmail = False
    				Exit Function
    			ElseIf lnOctect < 0 Then
    				IsEmail = False
    				Exit Function
    			End If
    			
    			' Look for private network settings
    			If lnIndex = 0 Then
    			
    				' Grab 2nd IP value
    				lnOctect2 = lsSubDomainArray(1)
    				
    				If Len(lnOctect2) > 3 Then
    					IsEmail = False
    					Exit Function
    				ElseIf Not IsNumeric(lnOctect2)Then
    					IsEmail = False
    					Exit Function
    				End If
    				
    				lnOctect2 = CInt(lnOctect2)
    				'	TCP/IP addresses reserved for 'private' networks are:
    				'				
    				'	10.0.0.0to 10.255.255.255				
    				'	172.16.0.0 to 172.31.255.255				
    				'	192.168.0.0to 192.168.255.255				
    				
    				Select Case lnOctect
    					
    					Case 10 ' Private Network
    						IsEmail = False
    						Exit Function
    					
    					Case 172
    						If lnOctect2 => 16 And lnOctect2 =< 31 Then
    							IsEmail = False
    							Exit Function
    						End If
    					
    					Case 192 ' Local Network
    						If lnOctect2 = 168 Then
    							IsEmail = False
    							Exit Function
    						End If
    					
    					Case 127 ' Local Machine
    						IsEmail = False
    						Exit Function				
    				
    				End Select
    			
    			End If
    			' End 'private' network check					
    		Else
    			
    			If lnIndex = UBound(lsSubDomainArray, 1) Then
    				
    				' Last domain can have 3 characters max
    				If Len(lsSubDomain) > 3 Then
    					IsEmail = False
    					Exit Function
    				ElseIf Not InStr(lsSubDomain, "-") = 0 Then
    					IsEmail = False
    					Exit Function
    				End If
    			
    			Else
    				' Domain, sub domain can only have 22 characters max
    				If Len(lsSubDomain) > 22 Then
    					IsEmail = False
    					Exit Function
    				End If
    			End If
    		End If
    	Next
    	
    	' Check for valid characters in username
    	lnStart = 1
    	Do While lnStart <= Len(lsUserName)
    		If InStr(lsDOMAIN_CHARACTERS, Mid(lsUserName, lnStart, 1)) Then
    			lnStart = lnStart + 1
    		Else
    			IsEmail = False
    			Exit Function
    		End If
    	Loop
    	
    	IsEmail = True
    End Function
    %>