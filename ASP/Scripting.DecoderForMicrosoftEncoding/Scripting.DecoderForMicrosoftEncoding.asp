<%
    '**************************************
    ' for :Scripting.Decoder for Microsoft E
    '     ncoding
    '**************************************
    'Copyright 2001. All rights reserved.
    'Lewis Moten
    'http://www.lewismoten.com
    'email: lewis@moten.com
    '**************************************
    ' Name: Scripting.Decoder for Microsoft 
    '     Encoding
    ' Description:This code lets you decode 
    '     your scripts. It comes in handy if you d
    '     eleted the original file or wrote over i
    '     t with the encoded version. Microsoft in
    '     cluded a program to encode, but choose n
    '     ot to let us decode with a scrdec progra
    '     m. More details can be found on my site 
    '     about this script including a live demon
    '     stration. Please leave comments here if 
    '     you wish to inquire about it.
    ' By: Lewis Moten
    '
    '
    ' Inputs:Encoded script string.
    '
    ' Returns:decoded script!
    '
    'Assumes:None
    '
    'Side Effects:unknown
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    ' THIS CODE AND INFORMATION IS PROVIDED 
    '     "AS IS" WITHOUT
    ' WARRANTY OF ANY KIND, EITHER EXPRESSED
    '     OR IMPLIED,
    ' INCLUDING BUT NOT LIMITED TO THE IMPLI
    '     ED WARRANTIES
    ' OF MERCHANTABILITY AND/OR FITNESS FOR 
    '     A PARTICULAR
    ' PURPOSE.
    ' Copyright 2000 - 2001. All rights rese
    '     rved.
    ' Lewis Moten
    ' http://www.lewismoten.com
    ' email: lewis@moten.com
    ' If you use this code on your site and 
    '     wish to be listed
    ' as one of the sites on my webpage that
    '     use my code, please
    ' email me the URL, Title, a short descr
    '     iption, and a 100x30 
    ' image button if you have one.
    ' I ask that you please help me promote 
    '     my site in
    ' exchange. This can be done simply by s
    '     ending a friend
    ' of yours an email, posting a link on y
    '     our site to mine,
    ' or even posting a message in a news gr
    '     oup, forum, IRC, etc.
    ' You can even do it by nominating it fo
    '     r sites that give
    ' out awards. If you link to my page, pl
    '     ease link to
    ' my home page at http://www.lewismoten.
    '     com, there is more
    ' to me then just programming. :)
    ' --------------------------------------
    '     ----------------------------------------
    '     
    Class clsScripting_Decoder
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Private mBytAsciiAry
    	Private mStrBase64
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Private Sub Class_Initialize()
    		mStrBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    		mBytAsciiAry = Array( _
    			&h00, &h00, &h00, &h01, &h01, &h01, &h02, &h02, &h02, &h03, &h03, &h03, &h04, &h04, &h04, &h05, &h05, &h05, _
    			&h06, &h06, &h06, &h07, &h07, &h07, &h08, &h08, &h08, &h57, &h6E, &h7B, &h00, &h00, &h00, &h0B, &h0B, &h0B, _
    			&h0C, &h0C, &h0C, &h00, &h00, &h00, &h0E, &h0E, &h0E, &h0F, &h0F, &h0F, &h10, &h10, &h10, &h11, &h11, &h11, _
    			&h12, &h12, &h12, &h13, &h13, &h13, &h14, &h14, &h14, &h15, &h15, &h15, &h16, &h16, &h16, &h17, &h17, &h17, _
    			&h18, &h18, &h18, &h19, &h19, &h19, &h1A, &h1A, &h1A, &h1B, &h1B, &h1B, &h1C, &h1C, &h1C, &h1D, &h1D, &h1D, _
    			&h1E, &h1E, &h1E, &h1F, &h1F, &h1F, &h2E, &h2D, &h32, &h47, &h75, &h30, &h7A, &h52, &h21, &h56, &h60, &h29, _
    			&h42, &h71, &h5B, &h6A, &h5E, &h38, &h2F, &h49, &h33, &h26, &h5C, &h3D, &h49, &h62, &h58, &h41, &h7D, &h3A, _
    			&h34, &h3E, &h35, &h32, &h36, &h65, &h5B, &h20, &h39, &h76, &h7C, &h5C, &h72, &h7A, &h56, &h43, &h7F, &h73, _
    			&h38, &h6B, &h66, &h39, &h63, &h4E, &h70, &h33, &h45, &h45, &h2B, &h6B, &h68, &h68, &h62, &h71, &h51, &h59, _
    			&h4F, &h66, &h78, &h09, &h76, &h5E, &h62, &h31, &h7D, &h44, &h64, &h4A, &h23, &h54, &h6D, &h75, &h43, &h71, _
    			&h00, &h00, &h00, &h7E, &h3A, &h60, &h00, &h00, &h00, &h5E, &h7E, &h53, &h40, &h00, &h40, &h77, &h45, &h42, _
    			&h4A, &h2C, &h27, &h61, &h2A, &h48, &h5D, &h74, &h72, &h22, &h27, &h75, &h4B, &h37, &h31, &h6F, &h44, &h37, _
    			&h4E, &h79, &h4D, &h3B, &h59, &h52, &h4C, &h2F, &h22, &h50, &h6F, &h54, &h67, &h26, &h6A, &h2A, &h72, &h47, _
    			&h7D, &h6A, &h64, &h74, &h39, &h2D, &h54, &h7B, &h20, &h2B, &h3F, &h7F, &h2D, &h38, &h2E, &h2C, &h77, &h4C, _
    			&h30, &h67, &h5D, &h6E, &h53, &h7E, &h6B, &h47, &h6C, &h66, &h34, &h6F, &h35, &h78, &h79, &h25, &h5D, &h74, _
    			&h21, &h30, &h43, &h64, &h23, &h26, &h4D, &h5A, &h76, &h52, &h5B, &h25, &h63, &h6C, &h24, &h3F, &h48, &h2B, _
    			&h7B, &h55, &h28, &h78, &h70, &h23, &h29, &h69, &h41, &h28, &h2E, &h34, &h73, &h4C, &h09, &h59, &h21, &h2A, _
    			&h33, &h24, &h44, &h7F, &h4E, &h3F, &h6D, &h50, &h77, &h55, &h09, &h3B, &h53, &h56, &h55, &h7C, &h73, &h69, _
    			&h3A, &h35, &h61, &h5F, &h61, &h63, &h65, &h4B, &h50, &h46, &h58, &h67, &h58, &h3B, &h51, &h31, &h57, &h49, _
    			&h69, &h22, &h4F, &h6C, &h6D, &h46, &h5A, &h4D, &h68, &h48, &h25, &h7C, &h27, &h28, &h36, &h5C, &h46, &h70, _
    			&h3D, &h4A, &h6E, &h24, &h32, &h7A, &h79, &h41, &h2F, &h37, &h3D, &h5F, &h60, &h5F, &h4B, &h51, &h4F, &h5A, _
    			&h20, &h42, &h2C, &h36, &h65, &h57, &h80, &h80, &h80, &h81, &h81, &h81, &h82, &h82, &h82, &h83, &h83, &h83, _
    			&h84, &h84, &h84, &h85, &h85, &h85, &h86, &h86, &h86, &h87, &h87, &h87, &h88, &h88, &h88, &h89, &h89, &h89, _
    			&h8A, &h8A, &h8A, &h8B, &h8B, &h8B, &h8C, &h8C, &h8C, &h8D, &h8D, &h8D, &h8E, &h8E, &h8E, &h8F, &h8F, &h8F, _
    			&h90, &h90, &h90, &h91, &h91, &h91, &h92, &h92, &h92, &h93, &h93, &h93, &h94, &h94, &h94, &h95, &h95, &h95, _
    			&h96, &h96, &h96, &h97, &h97, &h97, &h98, &h98, &h98, &h99, &h99, &h99, &h9A, &h9A, &h9A, &h9B, &h9B, &h9B, _
    			&h9C, &h9C, &h9C, &h9D, &h9D, &h9D, &h9E, &h9E, &h9E, &h9F, &h9F, &h9F, &hA0, &hA0, &hA0, &hA1, &hA1, &hA1, _
    			&hA2, &hA2, &hA2, &hA3, &hA3, &hA3, &hA4, &hA4, &hA4, &hA5, &hA5, &hA5, &hA6, &hA6, &hA6, &hA7, &hA7, &hA7, _
    			&hA8, &hA8, &hA8, &hA9, &hA9, &hA9, &hAA, &hAA, &hAA, &hAB, &hAB, &hAB, &hAC, &hAC, &hAC, &hAD, &hAD, &hAD, _
    			&hAE, &hAE, &hAE, &hAF, &hAF, &hAF, &hB0, &hB0, &hB0, &hB1, &hB1, &hB1, &hB2, &hB2, &hB2, &hB3, &hB3, &hB3, _
    			&hB4, &hB4, &hB4, &hB5, &hB5, &hB5, &hB6, &hB6, &hB6, &hB7, &hB7, &hB7, &hB8, &hB8, &hB8, &hB9, &hB9, &hB9, _
    			&hBA, &hBA, &hBA, &hBB, &hBB, &hBB, &hBC, &hBC, &hBC, &hBD, &hBD, &hBD, &hBE, &hBE, &hBE, &hBF, &hBF, &hBF, _
    			&hC0, &hC0, &hC0, &hC1, &hC1, &hC1, &hC2, &hC2, &hC2, &hC3, &hC3, &hC3, &hC4, &hC4, &hC4, &hC5, &hC5, &hC5, _
    			&hC6, &hC6, &hC6, &hC7, &hC7, &hC7, &hC8, &hC8, &hC8, &hC9, &hC9, &hC9, &hCA, &hCA, &hCA, &hCB, &hCB, &hCB, _
    			&hCC, &hCC, &hCC, &hCD, &hCD, &hCD, &hCE, &hCE, &hCE, &hCF, &hCF, &hCF, &hD0, &hD0, &hD0, &hD1, &hD1, &hD1, _
    			&hD2, &hD2, &hD2, &hD3, &hD3, &hD3, &hD4, &hD4, &hD4, &hD5, &hD5, &hD5, &hD6, &hD6, &hD6, &hD7, &hD7, &hD7, _
    			&hD8, &hD8, &hD8, &hD9, &hD9, &hD9, &hDA, &hDA, &hDA, &hDB, &hDB, &hDB, &hDC, &hDC, &hDC, &hDD, &hDD, &hDD, _
    			&hDE, &hDE, &hDE, &hDF, &hDF, &hDF, &hE0, &hE0, &hE0, &hE1, &hE1, &hE1, &hE2, &hE2, &hE2, &hE3, &hE3, &hE3, _
    			&hE4, &hE4, &hE4, &hE5, &hE5, &hE5, &hE6, &hE6, &hE6, &hE7, &hE7, &hE7, &hE8, &hE8, &hE8, &hE9, &hE9, &hE9, _
    			&hEA, &hEA, &hEA, &hEB, &hEB, &hEB, &hEC, &hEC, &hEC, &hED, &hED, &hED, &hEE, &hEE, &hEE, &hEF, &hEF, &hEF, _
    			&hF0, &hF0, &hF0, &hF1, &hF1, &hF1, &hF2, &hF2, &hF2, &hF3, &hF3, &hF3, &hF4, &hF4, &hF4, &hF5, &hF5, &hF5, _
    			&hF6, &hF6, &hF6, &hF7, &hF7, &hF7, &hF8, &hF8, &hF8, &hF9, &hF9, &hF9, &hFA, &hFA, &hFA, &hFB, &hFB, &hFB, _
    			&hFC, &hFC, &hFC, &hFD, &hFD, &hFD, &hFE, &hFE, &hFE, &hFF, &hFF, &hFF _
    		)
    	End Sub
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Public Function DecodeScriptFile(ByRef pStrExt, ByVal pStrScript, ByRef pLngTemp1, ByRef pStrTemp2)
    		Dim lStrEncodedScript
    		Dim lLngStart
    		Dim lLngEnd
    		Const lStrStartFlag = "#@~^"
    		Const lStrEndFlag = "^#~@"
    		
    		Do
    			lLngStart = InStr(1, pStrScript, lStrStartFlag)
    			If lLngStart = 0 Then Exit Do
    		
    			lLngEnd = InStr(lLngStart + 4, pStrScript, lStrEndFlag)
    			If lLngEnd = 0 Then Exit Do
    		
    			lStrEncodedScript = Mid(pStrScript, lLngStart, (lLngEnd + 4) - lLngStart)
    		
    			pStrScript = Left(pStrScript, lLngStart - 1) & Decode(lStrEncodedScript) & Mid(pStrScript, lLngEnd + 4)
    		Loop
    		
    		DecodeScriptFile = pStrScript
    	End Function
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Private Function Addition(ByRef pLngPosition)
    		Addition = CInt(Mid("0120121221210212021200122102122100212120200120210212001220012021", (pLngPosition Mod 64) + 1, 1))
    	End Function
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Private Function Decode(ByVal pStrScript)
    		Dim lLngLength
    		Dim lStrEncoded
    		Dim lLngPosition
    		Dim lLngCount
    		Dim lStrCharacter
    		Dim lStrEscapedCharacters
    		Dim lStrRealCharacters
    		Dim lStrDecoded
    		Dim lLngAddition
    		Dim lLngAscii
    		Dim lLngIndex
    		Dim lLngEscapePosition
    		Dim lStrEscapedCharacter
    		lStrEscapedCharacters = "!*$#&"
    		lStrRealCharacters = "<>@" & vbCrLf
    		
    		lLngLength = CheckLength(pStrScript)
    		 
    		pStrScript = Mid(pStrScript, 13)
    		pStrScript = Left(pStrScript, lLngLength)
    		lLngLength = Len(pStrScript)
    		
    		lLngCount = 0
    		For lLngPosition = 1 To lLngLength
    			lStrCharacter = Mid(pStrScript, lLngPosition, 1)
    			If lStrCharacter = "@" Then
    				lStrEscapedCharacter = Mid(pStrScript, lLngPosition + 1, 1)
    				lLngEscapePosition = InStr(1, lStrEscapedCharacters, lStrEscapedCharacter)
    				If lLngEscapePosition = 0 Then
    					lLngAddition = Addition(lLngPosition)
    					lLngAscii = Asc(lStrCharacter)
    					lLngIndex = (lLngAscii * 3) + lLngAddition
    					lStrDecoded = lStrDecoded & Chr(mBytAsciiAry(lLngIndex))
    				Else
    					lStrDecoded = lStrDecoded & Mid(lStrRealCharacters, lLngEscapePosition, 1)
    					lLngPosition = lLngPosition + 1
    				End If
    			Else
    				lLngAddition = Addition(lLngCount)
    				lLngAscii = Asc(lStrCharacter)
    				lLngIndex = (lLngAscii * 3) + lLngAddition
    				lStrDecoded = lStrDecoded & Chr(mBytAsciiAry(lLngIndex))
    			End If
    			lLngCount = lLngCount + 1
    		Next
    		
    		Decode = lStrDecoded
    	End Function
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Private Function CheckLength(ByRef pStrScript)
    		Dim lStrLength
    		Dim lLngIndex
    		Dim lLngValue
    		Dim lLngTotal
    		lStrLength = Mid(pStrScript, 5, 4)
    		lStrLength = Base64Decode(lStrLength)
    		
    		For lLngIndex = 1 To Len(lStrLength)
    			lLngValue = Asc(Mid(lStrLength, lLngIndex, 1))
    			lLngTotal = lLngTotal + lLngValue * (255 ^ (lLngIndex -1))
    		next
    		CheckLength = lLngTotal
    	End Function
    ' --------------------------------------
    '     ----------------------------------------
    '     
    	Function Base64decode(ByVal pStrContents)
    	Dim lStrResult
    	Dim lLngPosition
    	Dim lStrGroup64
    	Dim lStrGroupBinary
    	Dim lStrChar1
    	Dim lStrChar2
    	Dim lStrChar3
    	Dim lStrChar4
    	Dim lByt1
    	Dim lByt2
    	Dim lByt3
    	
    	If Len(pStrContents) Mod 4 > 0 Then
    			pStrContents = pStrContents & String(4 - (Len(pStrContents) Mod 4), "=")
    		End If
    		
    	lStrResult = ""
    	
    	For lLngPosition = 1 To Len(pStrContents) Step 4
    	lStrGroupBinary = ""
    	lStrGroup64 = Mid(pStrContents, lLngPosition, 4)
    	lStrChar1 = InStr(mStrBase64, Mid(lStrGroup64, 1, 1)) - 1
    	lStrChar2 = InStr(mStrBase64, Mid(lStrGroup64, 2, 1)) - 1
    	lStrChar3 = InStr(mStrBase64, Mid(lStrGroup64, 3, 1)) - 1
    	lStrChar4 = InStr(mStrBase64, Mid(lStrGroup64, 4, 1)) - 1
    	lByt1 = Chr(((lStrChar2 And 48) \ 16) Or (lStrChar1 * 4) And &HFF)
    	lByt2 = lStrGroupBinary & Chr(((lStrChar3 And 60) \ 4) Or (lStrChar2 * 16) And &HFF)
    	lByt3 = Chr((((lStrChar3 And 3) * 64) And &HFF) Or (lStrChar4 And 63))
    	lStrGroupBinary = lByt1 & lByt2 & lByt3
    	
    	lStrResult = lStrResult + lStrGroupBinary
    	Next
    	Base64decode = lStrResult
    	End Function
    ' --------------------------------------
    '     ----------------------------------------
    '     
    End Class
    ' --------------------------------------
    '     ----------------------------------------
    '     
    %>