    <%
    '**************************************
    ' for :HTMLDecode
    '**************************************
    'Copyright (C) 2001, Lewis Edward Moten III. All rights reserved.
    '**************************************
    ' Name: HTMLDecode
    ' Description:This function decodes HTML
    '     Encoded strings. I found that when some 
    '     browsers post data, it converts double-b
    '     yte characters to html encoded character
    '     s. This script searches for all numerica
    '     l entities as well as a few popular name
    '     d entities. You will need to tell the se
    '     rver and the web browser that you are wo
    '     rking with Unicode characters. The code 
    '     has been included for working with the U
    '     TF-7 character set. Enjoy!
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
    
    ' These next 3 lines are needed.
    ' They tell server to use Unicode 7 byte
    '     character set.
    ' You may wish to place them at the top 
    '     of your ASP Pages
    Session.CodePage = 65000
    Response.AddHeader "Content-Type", "text/html; CHARSET=utf-7"
    Response.Write("<META http-equiv=""Content-Type"" content=""text/html; CHARSET=utf-7"">")
    ' Demonstration ...
    Response.Write HTMLDecode("He said &#8220;Hello Joe&#8221; and &#8211; he said &#9827; &#8220;morning.&#8221; &Alpha;" & ChrW(9827))
    Function HTMLDecode(ByVal pstrHTML)
    	' This function Decodes HTML encoded characters.
    	
    	Dim lobjRegExp		' Regular Expression
    	Dim lobjMatches		' Collection of matches
    	Dim lobjMatch		' Single match
    	Dim lstrEntityAry	' Array of HTML entities
    	Dim lstrEntity		' HTML entity
    	Dim llngCharCode	' Character Code for HTML Entity
    	Dim llngIndex		' Index within entity array
    	Dim llngMaxIndex	' Maximum Index within entity array
    	
    	' Make sure we have data
    	If Trim(pstrHTML & "") = "" Then Exit Function
    	
    	' Regular Expression
    	Set lobjRegExp = New RegExp
    	
    	' Setup properties to find HTML encoded characters with digits
    	lobjRegExp.Global = True
    	lobjRegExp.IgnoreCase = False
    	lobjRegExp.Multiline = True
    	lobjRegExp.Pattern = "&#(\d+);"
    	
    	' Grab Matches
    	Set lobjMatches = lobjRegExp.Execute(pstrHTML)
    	
    	' Loop through matches
    	For Each lobjMatch In lobjMatches
    		pstrHTML = Replace(pstrHTML, lobjMatch.value, ChrW(lobjMatch.SubMatches(0)))
    	Next
    	
    	' Clean Up
    	Set lobjMatch = Nothing
    	Set lobjMatches = Nothing
    	Set lobjRegExp = Nothing
    	
    	' Split string into an array.
    	' Case-Sensitive Format is
    		' Character Code, Entity, Character Code, Entity, ...
    	lstrEntityAry = Split( _
    		"402,fnof,913,Alpha,914,Beta,915,Gamma,916,Delta,917,Epsilon," & _
    		"918,Zeta,919,Eta,920,Theta,921,Iota,922,Kappa,923,Lambda," & _
    		"924,Mu,925,Nu,926,Xi,927,Omicron,928,Pi,929,Rho,931,Sigma," & _
    		"932,Tau,933,Upsilon,934,Phi,935,Chi,936,Psi,937,Omega,945,alpha," & _
    		"946,beta,947,gamma,948,delta,949,epsilon,950,zeta,951,eta," & _
    		"952,theta,953,iota,954,kappa,955,lambda,956,mu,957,nu,958,xi," & _
    		"959,omicron,960,pi,961,rho,962,sigmaf,963,sigma,964,tau," & _
    		"965,upsilon,966,phi,967,chi,968,psi,969,ometa,977,thetasym," & _
    		"978,upsih,982,piv,8226,bull,8230,hellip,8242,prime,8243,Prime," & _
    		"8254,oline,8260,frasl,8472,weierp,8465,image,8476,real," & _
    		"8482,trade,8501,alefsym,8592,larr,8593,uarr,8594,rarr," & _
    		"8595,darr,8596,harr,8629,crarr,8656,lArr,8657,uArr,8658,rArr," & _
    		"8659,dArr,8660,hArr,8704,forall,8706,part,8707,exist,8709,empty," & _
    		"8711,nabla,8712,isin,8713,notin,8715,ni,8719,prod,8722,sum," & _
    		"8722,minus,8727,lowast,8730,radic,8733,prop,8734,infin,8736,ang," & _
    		"8869,and,8870,or,8745,cap,8746,cup,8747,int,8756,there4,8764,sim," & _
    		"8773,cong,8773,asymp,8800,ne,8801,equiv,8804,le,8805,ge,8834,sub," & _
    		"8835,sup,8836,nsub,8838,sube,8839,supe,8853,oplus,8855,otimes," & _
    		"8869,perp,8901,sdot,8968,lceil,8969,rceil,8970,lfloor,8971,rfloor," & _
    		"9001,lang,9002,rang,9674,loz,9824,spades,9827,clubs,9829,hearts," & _
    		"9830,diams", ",")
    	
    	' How many indexes do we have?
    	llngMaxIndex = UBound(lstrEntityAry)
    	
    	' Loop through each HTML Entity and convert it to a character
    	For llngIndex = 0 To llngMaxIndex Step 2 ' (two indexes = 1 record)
    		llngCharCode = lstrEntityAry(llngIndex)
    		lstrEntity = "&" & lstrEntityAry(llngIndex + 1) & ";"
    		
    		pstrHTML = Replace(pstrHTML, lstrEntity, ChrW(llngCharCode))
    	
    	Next
    			
    	' Replace encoded non-digits.
    	pstrHTML = Replace(pstrHTML, "&quot;", """")
    	pstrHTML = Replace(pstrHTML, "&lt;", "<")
    	pstrHTML = Replace(pstrHTML, "&gt;", "<")
    	pstrHTML = Replace(pstrHTML, "&amp;", "<")
    	
    	' Return the results
    	HTMLDecode = pstrHTML
    End Function
    %>