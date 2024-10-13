    <%
    '**************************************
    ' Name: Lewies URLDecode
    ' Description:Xao Xiong originally poste
    '     d his version 
    'of DecodeURL and asked for a reply for 
    'a better way. I replied, but it didn't 
    'turn out too well, so i'm posting it here 
    'for him to see.
    'The approach i took is that when you look 
    'at an encoded URL string, 2 hex characters 
    'follow a percent sign. XAO listed a few 
    'to translate into actual characters, but 
    'he didn't account for all 256 of them. 
    'Check it out!
    'See http://www.planetsourcecode.com/xq/ASP/
    'txtCodeId.6728/lngWId.4/qx/vb/scripts/
    'ShowCode.htm
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
    
    Response.Write DecodeURL("http://www.lewismoten.com/1%2B2%3D3.asp")
    Function DecodeURL(ByRef pstrURL)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrChar
    	Dim lstrResult
    	
    	llngMaxIndex = Len(pstrURL)
    	llngIndex = 1
    	
    	Do While llngIndex <= llngMaxIndex
    		lstrChar = Mid(pstrURL, llngIndex, 1)
    		If lstrChar = "%" Then
    			lstrResult = lstrResult & chr("&h" & Mid(pstrURL, llngIndex + 1, 2))
    			llngIndex = llngIndex + 3
    		ElseIf lstrChar = "+" Then
    			lstrResult = lstrResult & " "
    			llngIndex = llngIndex + 1
    		Else
    			lstrResult = lstrResult & lstrChar
    			llngIndex = llngIndex + 1
    		End If
    	Loop
    	
    	DecodeURL = lstrResult
    End Function
    %>