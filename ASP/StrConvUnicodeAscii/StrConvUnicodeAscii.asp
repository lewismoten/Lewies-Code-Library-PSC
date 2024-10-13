<%
    '**************************************
    ' Name: StrConv Unicode/Ascii
    ' Description:Converts between unicode a
    '     nd ANSI using ADODB Streams and recordse
    '     ts. I was pressing myself to the limit t
    '     o find a faster way to convert strings b
    '     ack and forth between unicode and ANSI d
    '     uring file uploads. This version goes th
    '     e distance and has great speed.
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
    
    Const vbFromUnicode = 128
    Const vbUnicode = 64
    Public Function StrConv(ByRef stringData, ByRef conversion)
    	Dim Stream
    	Set Stream = Server.CreateObject("ADODB.Stream")
    	' Charsets	
    	'	Windows-1252
    	'	Windows-1257
    	'	UTF-8
    	'	UTF-7
    	'	ASCII
    	'	X-ANSI
    	Const UnicodeCharaset = "Windows-1252"
    	Const BinaryCharset = "X-ANSI"
    	Select Case conversion
    		Case vbFromUnicode
    			' Converts a Unicode string to Ascii
    			With Stream
    				.Charset = UnicodeCharaset
    				.Type = adTypeText
    				.Open
    				.WriteText stringData
    				.Position = 0
    				.Charset = BinaryCharset
    				.Type = adTypeBinary
    				StrConv = MidB(.Read, 1)
    			End With
    			
    		Case vbUnicode
    			' Converts an Ascii string to Unicode
    			Dim Length
    			Dim Buffer
    			
    			If TypeName(stringData) = "Null" Then
    				CStrU = ""
    				Exit Function
    			End If
    			
    			stringData = MidB(stringData, 1)
    		
    			Length = LenB(stringData)
    			Dim Rs
    			Set Rs = Server.CreateObject("ADODB.Recordset")
    			Call Rs.Fields.Append("BinaryData", adLongVarBinary, Length)
    			Rs.Open
    			Rs.AddNew
    			Rs.Fields("BinaryData").AppendChunk(stringData & ChrB(0))
    			Rs.Update
    			Buffer = Rs.Fields("BinaryData").GetChunk(Length)
    			Rs.Close
    			Set Rs = Nothing
    			With Stream
    				.Charset = BinaryCharset
    				.Type = adTypeBinary
    				.Open
    				Call .Write(Buffer)
    				.Position = 0
    				.Type = adTypeText
    				.Charset = UnicodeCharaset
    			End With
    			
    			StrConv = Stream.ReadText(-1)
    	End Select
    	Stream.Close
    	Set Stream = Nothing
    End Function
    %>