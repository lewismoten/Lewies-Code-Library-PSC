<%
    '**************************************
    ' for :Convert StringIP to NumericIP
    '**************************************
    'Copyright (C) 1999, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Convert StringIP to NumericIP
    ' Description:Converts a string ip addre
    '     ss ("192.168.0.1") to a Long number (323
    '     2235521). One of the resons to do this w
    '     ould be to store IP addresses in databas
    '     es. Numbers greatly reduce the size requ
    '     ired to store this information.
    ' By: Lewis Moten
    '
    '
    ' Inputs:asNewIP - String IP address ("1
    '     92.168.0.1") to convert to a number.
    '
    ' Returns:Returns a Long Integer represe
    '     nting the IP address (3232235521)
    '
    'Assumes:This function assumes that your
    '     IP address has 4 integers delimited by d
    '     ecimals and that the numbers range from 
    '     0 to 255.
    '
    'Side Effects:If storing number within a
    '     n Access Database, The variable type "Lo
    '     ng" ranges from -2147483648 to 214748364
    '     7. You will need to subtract 2147483648 
    '     from the number before querying the data
    '     base.
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Function CLngIP(ByVal asNewIP)
    	Dim lnResults
    	Dim lnIndex
    	Dim lnIpAry
    	' Split the IP address using the dot as a delimiter
    	lnIpAry = Split(asNewIP, ".", 4)
    	' Loop through each number in the IP address
    	For lnIndex = 0 To 3
    		
    		' If we are not working with the last number...
    		If Not lnIndex = 3 Then
    			
    			' Convert the number to a value range that can be parsed from the others
    			lnIpAry(lnIndex) = lnIpAry(lnIndex) * (256 ^ (3 - lnIndex))
    		
    		End If
    		
    		' Add the number to the results
    		lnResults = lnResults + lnIpAry(lnIndex)
    	
    	Next
    	' If storing number within an Access Database,
    	' The variable type "Long" ranges from -2147483648 to 2147483647
    	' You will need to subtract 2147483648 from the number
    	' before querying the database.
    	
    	' lnResults = lnResults - 2147483648
    	' Return the results
    	CLngIP = lnResults
    End Function

%>