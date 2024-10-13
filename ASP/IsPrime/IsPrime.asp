<%
    '**************************************
    ' Name: IsPrime
    ' Description:Checks to see if a number 
    '     is a primary number. Prime numbers are c
    '     ommonly found in encryption schemes that
    '     use Public/Private keys.
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
    
    Function IsPrime(ByRef pLngNumber)
    	Dim lLngSquare
    	Dim lLngIndex
    	IsPrime = False
    If pLngNumber < 2 Then Exit Function
    If pLngNumber Mod 2 = 0 Then Exit Function
    lLngSquare = Sqr(pLngNumber)
    For lLngIndex = 3 To lLngSquare Step 2
    If pLngNumber Mod lLngIndex = 0 Then Exit Function
    Next
    	IsPrime = True
    End Function

%>