<%
  '**************************************
    ' Name: ANSI to Unicode
    ' Description:Converts from ANSI to Unic
    '     ode very fast. Inspired by code found in
    '     UltraFastAspUpload by Cakkie (on PSC). T
    '     his should work slightly faster then Cak
    '     kies due to how some of the code has bee
    '     n arranged.
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
    
    ' make a sample binary string of 2000 ra
    '     ndom letters
    Dim test
    Dim c
    For i = 1 to 2000
    	c = 65 + Int(Rnd * 25)
    	test = test & ChrB(AscB(Chr(c)))
    	If i Mod 4 = 0 Then
    		test = test & ChrB(AscB(" "))
    	End If
    Next
    Response.Write ANSIToUnicode(test)
    Function ANSIToUnicode(ByRef pbinBinaryData)
    	Dim lbinData	' Binary Data (ANSI)
    	Dim llngLength	' Length of binary data (byte count)
    	Dim lobjRs		' Recordset
    	Dim lstrData	' Unicode Data
    	' VarType Reference
    	'8 = Integer (this is expected var type)
    	'17 = Byte Subtype
    	' 8192 = Array
    	' 8209 = Byte Subtype + Array
    	
    	Set lobjRs = Server.CreateObject("ADODB.Recordset")
    	If VarType(pbinBinaryData) = 8 Then
    		' Convert integers(4 bytes) to Byte Subtype Array (1 byte)
    		llngLength = LenB(pbinBinaryData)
    		If llngLength = 0 Then
    			lbinData = ChrB(0)
    		Else
    			Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
    			Call lobjRs.Open()
    			Call lobjRs.AddNew()
    			Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0)) ' + Null terminator
    			Call lobjRs.Update()
    			lbinData = lobjRs.Fields("BinaryData").GetChunk(llngLength)
    			Call lobjRs.Close()
    		End If
    	Else
    		lbinData = pbinBinaryData
    	End If
    	' Do REAL conversion now!	
    	llngLength = LenB(lbinData)
    	If llngLength = 0 Then
    		lstrData = ""
    	Else
    		Call lobjRs.Fields.Append("BinaryData", adLongVarChar, llngLength)
    		Call lobjRs.Open()
    		Call lobjRs.AddNew()
    		Call lobjRs.Fields("BinaryData").AppendChunk(lbinData)
    		Call lobjRs.Update()
    		lstrData = lobjRs.Fields("BinaryData").Value
    		Call lobjRs.Close()
    	End If
    				
    	Set lobjRs = Nothing
    	ANSIToUnicode = lstrData
    	
    End Function
    %>
