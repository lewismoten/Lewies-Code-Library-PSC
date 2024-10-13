<%
    '**************************************
    ' for :Convert LongIP to StringIP
    '**************************************
    'Copyright (c) by Lewis Moten, All rights reserved.
    '**************************************
    ' Name: Convert LongIP to StringIP
    ' Description:This function converts a L
    '     ong Number(3232235521) into an IP Addres
    '     s ("192.168.0.1").
    ' By: Lewis Moten
    '
    '
    ' Inputs:anNewIP - IP Address as a Long 
    '     Number (no dots)
    '
    ' Returns:Returns the string representat
    '     ion of an IP Address ("192.168.0.1")
    '
    'Assumes:If pulling number from an Acces
    '     s Database, The variable type "Long" ran
    '     ges from -2147483648 to 2147483647. You 
    '     will first need to add 2147483648 to the
    '     number to parse correctly.
    '
    'Side Effects:Assumes that the number is
    '     a positive whole number that is not grea
    '     ter then 4294967295
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Function CStrIP(ByVal anNewIP)
    	Dim lsResults		' Results to be returned
    	Dim lnTemp			' Temporary value being parsed
    	Dim lnIndex			' Position of number being parsed
    	' If pulling number from an Access Database,
    	' The variable type "Long" ranges from -2147483648 to 2147483647
    	' You will first need to add 2147483648 to the number to parse correctly.
    	' anNewIP = anNewIP + 2147483648
    	' Parse highest numbers first
    	For lnIndex = 3 To 0 Step -1
    	
    		' Parse the current value for this position
    		lnTemp = Int(anNewIP / (256 ^ lnIndex))
    		
    		' Append the number to the final results delimited by a dot
    		lsResults = lsResults & lnTemp & "."
    		
    		' Remove the number that we just parsed
    		anNewIP = anNewIP - (lnTemp * (256 ^ lnIndex))
    		
    	Next
    	
    	' Cut off last dot
    	lsResults = Left(lsResults, Len(lsResults) - 1)
    	
    	' Return the results
    	CStrIP = lnResults
    	
    End Function

%>