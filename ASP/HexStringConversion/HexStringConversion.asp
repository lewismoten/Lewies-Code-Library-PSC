<%
    '**************************************
    ' for :Hex/String Conversion
    '**************************************
    'Copyright (c) 2001, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Hex/String Conversion
    ' Description:These functions allow you 
    '     to easily convert a string to its hex va
    '     lue and back again.
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
    
    email = "lewis@moten.com"
    Response.Write "Original value: " & email & "<BR>"
    email = StringToHex(email)
    Response.Write "Hex value: " & email & "<BR>"
    email = HexToString(email)
    Response.Write "Back to string value: " & email & "<BR>"
    Function StringToHex(ByRef pstrString)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrHex
    	llngMaxIndex = Len(pstrString)
    	For llngIndex = 1 To llngMaxIndex
    		lstrHex = lstrHex & Right("0" & Hex(Asc(Mid(pstrString, llngIndex, 1))), 2)
    	Next
    	StringToHex = lstrHex
    End Function
    Function HexToString(ByRef pstrHex)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrString
    	llngMaxIndex = Len(pstrHex)
    	For llngIndex = 1 To llngMaxIndex Step 2
    		lstrString = lstrString & Chr("&h" & Mid(pstrHex, llngIndex, 2))
    	Next
    	HexToString = lstrString
    End Function
%>