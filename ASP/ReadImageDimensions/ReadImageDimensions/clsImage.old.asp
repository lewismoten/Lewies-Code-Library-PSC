<%
Class clsImage

	Private mStrBinaryData
	Private mLngWidth
	Private mLngHeight
	Private mStrType
	Private mStrContentType
	Private mLngSize
	Private mStrPath
	
	Private Sub Class_Initialize()
		mStrBinaryData = ChrB(0)
		mLngWidth = -1
		mLngHeight = -1
		mLngSize = -1
		mStrPath = "Undefined"
		mStrType = "Unknown"
		mStrContentType = "application/octet-stream"
	End Sub
	
	Public Sub Read(ByVal pStrFilePath)
		
		' Reset
		mStrBinaryData = ""
		mLngWidth = -1
		mLngHeight = -1
		mLngSize = -1
		mStrType = "Unknown"
		mStrContentType = "application/octet-stream"
		
		If InStr(1, pStrFilePath, ":\") = 0 Then
			pStrFilePath = Server.MapPath(pStrFilePath)
		End If
		
		mStrPath = pStrFilePath
		
		Dim lObjFSO
		Dim lObjFile
		Set lObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		If lObjFSO.FileExists(pStrFilePath) Then
			Set lObjFile = lObjFSO.OpenTextFile(pStrFilePath)
			If Not lObjFile.AtEndOfStream Then
				mStrBinaryData = ChrB(Asc(lObjFile.Read(1)))
				While Not lObjFile.AtEndOfStream
					 mStrBinaryData = mStrBinaryData & ChrB(Asc(lObjFile.Read(1)))
				Wend
			End If
			lObjFile.Close
			Call ReadDimensions()
		End If
		
		Set lObjFSO = Nothing
		
	End Sub
	
	Public Property Let DataStream(ByRef pStrBinaryData)
		mStrPath = "DataStream"
		mStrBinaryData = pStrBinaryData
		Call ReadDimensions()
	End Property
	
	Public Property Get DataStream()
		DataStream = mStrBinaryData
	End Property
	
	Public Property Get Width()
		Width = mLngWidth
	End Property
	
	Public Property Get Height()
		Height = mLngHeight
	End Property
	
	Public Property Get ImageType()
		ImageType = mStrType
	End Property
	
	Public Property Get ContentType()
		ContentType = mStrContentType
	End Property
	
	Public Property Get Size()
		Size = mLngSize
	End Property
	
	Public Property Get Path()
		Path = mStrPath
	End Property
	
	Private Sub ReadDimensions() 
		
		mLngWidth = -1
		mLngHeight = -1
		mLngSize = LenB(mStrBinaryData)
		mStrType = "Unknown"
		mStrContentType = "application/octet-stream"
		
		' I refer to Ascii data as Binary data or "BIN" in this script.
		
		Dim lBinGIF		' Signature of GIF
		Dim lBinJPG		' Signature of JPG
		Dim lBinBMP		' Signature of BMP
		Dim lBinPNG		' Signature of PNG
		
		lBinGIF = ChrB(Asc("G")) & ChrB(Asc("I")) & ChrB(Asc("F"))
		lBinJPG = ChrB(Asc("J")) & ChrB(Asc("F")) & ChrB(Asc("I")) & ChrB(Asc("F"))
		lBinBMP = ChrB(Asc("B")) & ChrB(Asc("M"))
		lBinPNG = ChrB(&h89) & ChrB(Asc("P")) & ChrB(Asc("N")) & ChrB(Asc("G"))
		
		' GIF File
		If InStrB(1, mStrBinaryData, lBinGIF) = 1 Then
			mStrType = "GIF"
			mStrContentType = "image/gif"
			
			mLngWidth = CLng("&h" & HexAt(8) & HexAt(7))
			mLngHeight = CLng("&h" & HexAt(10) & HexAt(9))
		' JPEG file
		ElseIf InStrB(1, mStrBinaryData, lBinJPG) = 7 Then
			mStrType = "JPG"
			mStrContentType = "image/jpeg"
			
			' Prefix found before image dimensions		
			lBinPrefix = ChrB(&h00) & ChrB(&h11) & ChrB(&h08)

			' Find the last prefix (so we don't confuse it with data)		
			lLngStart = 1
			Do
				If InStrB(lLngStart, mStrBinaryData, lBinPrefix) + 3 = 3 Then Exit Do
				lLngStart = InStrB(lLngStart, mStrBinaryData, lBinPrefix) + 3
			Loop
			
			' If a prefix was found
			If Not lLngStart = 1 Then
				mLngWidth = CLng("&h" & HexAt(lngStart+2) & HexAt(lngStart+3))
				mLngHeight = CLng("&h" & HexAt(lngStart) & HexAt(lngStart+1))
			End If

		' Bitmap File
		ElseIf InStrB(1, mStrBinaryData, lBinBMP) = 1 Then
			mStrType = "BMP"
			mStrContentType = "image/bmp"
			mLngWidth = CLng("&h" & HexAt(22) & HexAt(21) & HexAt(20) & HexAt(19))
			mLngHeight = CLng("&h" & HexAt(26) & HexAt(25) & HexAt(24) & HexAt(23))
		' PNG File
		ElseIf InStrB(1, mStrBinaryData, lBinPNG) = 1 Then
			mStrType = "PNG"
			mStrContentType = "image/png"
			mLngWidth = CLng("&h" & HexAt(17) & HexAt(18) & HexAt(19) & HexAt(20))
			mLngHeight = CLng("&h" & HexAt(21) & HexAt(22) & HexAt(23) & HexAt(24))
		End If

	End Sub
	Private Function HexAt(ByRef pLngPosition)
		If pLngPosition > LenB(mStrBinaryData) Or pLngPosition <= 0 Then Exit Function
		HexAt = Right("0" & Hex(AscB(MidB(mStrBinaryData, pLngPosition, 1))), 2)
	End Function
End Class
%>