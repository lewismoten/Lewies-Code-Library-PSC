<%
    '**************************************
    ' Name: encrypt4.asp
    ' Description:Encrypts text using the en
    '     igma algotithm used in World War II. Thi
    '     s security is easily breakable, but peop
    '     le seem to be using this. The next step 
    '     is XOR encryption - which is also easily
    '     breakable. Use this only to get an idea 
    '     behind how security algorithms have evol
    '     ved over time and identify its weakeness
    '     es. This code is an improvement in clari
    '     ty and support of characters such as NUL
    '     Ls to encrypt2.asp by Barry Beatti - htt
    '     p://www.planet-source-code.com/vb/script
    '     s/ShowCode.asp?lngWId=4&txtCodeId=7211
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
    ' --------------------------------------
    '     --------------------------------
    Function EnigmaEncrypt(Message, Password)
    	Dim MessageLength
    	Dim PasswordLength
    	Dim Index
    	Dim Ascii
    	Dim Key
    	Dim NewAscii
    	Dim Encrypted
    	If Message = "" Or Password = "" Then 
    		EnigmaEncrypt = Message
    		Exit Function
    	End If
    	' Get lengths of message and password
    	MessageLength = Len(Message)
    	PasswordLength = Len(Password)
    	
    	' loop through each character in message
    	For Index = 1 To MessageLength
    		' Get ascii value of message character
    		Ascii = Asc(Mid(Message, Index, 1))
    		
    		' Get ascii value of key at correct position
    		Key = Asc(Mid(Password, (Index Mod PasswordLength + 1), 1))
    		
    		' Add key to ascii
    		NewAscii = Ascii + Key
    		
    		' Working with bytes, loop to low numbers if below zero
    		If NewAscii > 255 Then NewAscii = NewsAscii - 255
    		
    		' apend ascii as a double hex character to encrypted message
    		Encrypted = Encrypted & Right("0" & Hex(NewAscii), 2)
    	Next
    	' Return the encrypted message	
    	EnigmaEncrypt = Encrypted
    End Function
    ' --------------------------------------
    '     --------------------------------
    Function EnigmaDecrypt(Message, Password)
    	Dim MessageLength
    	Dim PasswordLength
    	Dim Index
    	Dim Ascii
    	Dim Key
    	Dim NewAscii
    	Dim Decrypted
    	' Make certain that a message and password exists
    	If Message = "" Or Password = "" Then 
    		EnigmaDecrypt = Message
    		Exit Function
    	End If
    	
    	' Each character is made of two hex characters, so divide message by two
    	MessageLength = Len(Message) \ 2
    	PasswordLength = Len(Password)
    	
    	' Loop through each letter
    	For Index = 1 To MessageLength
    		' Get ascii value of hex
    		Ascii = CInt("&h" & Mid(Message, ((Index -1) * 2) + 1, 2))
    		
    		' get ascii value of key at correct position
    		Key = Asc(Mid(Password, (Index Mod PasswordLength + 1), 1))
    		
    		' Subtract Key value
    		NewAscii = Ascii - Key
    		' Working with bytes, loop to high numbers if below zero
    		If NewAscii < 0 Then NewAscii = NewsAscii + 255
    		
    		' Append to decrypted value
    		Decrypted = Decrypted & Chr(NewAscii)
    	Next
    	
    	' Return decrypted message
    	EnigmaDecrypt = Decrypted
    End Function
    ' --------------------------------------
    '     --------------------------------
    Dim Encrypt
    Dim Message
    Dim Password
    If Request.QueryString("Encrypt") = "" Then
    	Encrypt = True
    Else
    	Password = Request.QueryString("Password")
    	Message = Request.QueryString("Message")
    	Encrypt = Request.QueryString("Encrypt") = "True"
    	
    	If Encrypt Then
    		Message = EnigmaEncrypt(Message, Password)
    	Else
    		Message = EnigmaDecrypt(Message, Password)
    	End If
    	Encrypt = Not Encrypt
    End If
    %>
    <FORM>
    	Password: <INPUT type="text" name="Password" value="<%=Server.HTMLEncode(Password)%>"><BR>
    	<BR>
    	Message: <INPUT type="text" name="Message" value="<%=Server.HTMLEncode(Message)%>"><BR>
    	<BR>
    	<INPUT type="radio" name="Encrypt" value="True"<%If Encrypt Then Response.Write(" checked")%>>Encrypt<BR>
    	<INPUT type="radio" name="Encrypt" value="False"<%If Not Encrypt Then Response.Write(" checked")%>>Decrypt<BR>
    	<BR>
    	<INPUT type="Submit">
    </FORM>

