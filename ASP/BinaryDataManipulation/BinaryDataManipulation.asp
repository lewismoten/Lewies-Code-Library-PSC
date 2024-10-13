<%
    '**************************************
    ' Name: Binary Data Manipulation
    ' Description:These two functions allow 
    '     you to convert between Unicode and Ascii
    '     strings. This is great if you are workin
    '     g with the Request.BinaryRead/BinaryWrit
    '     e methods or binary data within a databa
    '     se.
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
    
    Function AsciiToUnicode(ByRef pstrAscii)
    	
    	Dim llngLength
    	Dim llngIndex
    	Dim llngAscii
    	Dim lstrUnicode
    	
    	llngLength = LenB(pstrAscii)
    	
    	For llngIndex = 1 to llngLength
    		llngAscii = AscB(MidB(pstrAscii, llngIndex, 1))
    		lstrUnicode = lstrUnicode & Chr(llngAscii)
    	Next
    	
    	AsciiToUnicode = lstrUnicode
    	
    End Function
    Function UnicodeToAscii(ByRef pstrUnicode)
    	Dim llngLength
    	Dim llngIndex
    	Dim llngAscii
    	Dim lstrAscii
    	
    	llngLength = Len(pstrUnicode)
    	
    	For llngIndex = 1 to llngLength
    		llngAscii = Asc(Mid(pstrUnicode, llngIndex, 1))
    		lstrAscii = lstrUnicode & ChrB(llngAscii)
    	Next
    	
    	UnicodeToAscii = lstrAscii
    End Function

%>