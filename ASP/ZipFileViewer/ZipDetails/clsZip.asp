<%
Class clsZip
	Private mbin_Zip
	Private mobj_Files()
	Private mlng_Files
	
	Sub ZipLoad(pstrFileName)
		Dim lobjFSO
		Dim llngTristateFalse
		Dim llngForReading
		Dim lobjFile
		
		llngTristateFalse = 0
		llngForReading = 1

		mbin_Zip = ""
		
		If pstrFileName = "" Then Exit Sub

		If InStr(1, pstrFileName, ":\") = 0 Then
			pstrFileName = Server.MapPath(pstrFileName)
		End If
		Set lobjFSO = Server.CreateObject("Scripting.FileSystemObject")

		If lobjFSO.FileExists(pstrFileName) Then

			Set lobjFile = lobjFSO.OpenTextFile(pstrFileName, llngForReading, False, llngTristateFalse)

			While Not lobjFile.AtEndOfStream
				mbin_Zip = mbin_Zip & ChrB(Asc(lobjFile.Read(1)))
			Wend
			lobjFile.Close
			Set lobjFile = Nothing
				
		End If
			
		Set lobjFSO = Nothing
			
		Call ParseZips()

	End Sub
	
	Public Property Let ZipData(ByRef pbinBinaryData)
		mbin_Zip = pbinBinaryData
		Call ParseZips()
	End Property
	Public Property Get FileCount()
		FileCount = mlng_Files
	End Property
	Public Property Get GetFile(ByRef plngIndex)
		Set GetFile = mobj_Files(plngIndex-1)
	End Property
'===========================================================================================	
	Private Sub ParseZips()
		Dim llngOffSet
		mlng_Files = 0
		llngOffSet = 0
		
		If LenB(mbin_Zip) = 0 Then Exit Sub
		
		Do
			'If GetNumber(llngOffset + 19, 4) = 0 Then Exit Do

			' If name is not specified, then we are looking at
			' a folder, or overall view of zip file

			' Find next PK 3.04 record
			llngOffset = InStrB(llngOffset + 1, mbin_zip, ChrB(&h50) & ChrB(&h4B) & ChrB(&h03) & ChrB(&h04))
			If llngOffset = 0 Then Exit Do
			llngOffset = llngOffset - 1
			
			ReDim Preserve mobj_Files(mlng_Files)
		
			Set mobj_Files(mlng_Files) = New clsZipFile
			With mobj_Files(mlng_Files)
				.Signature = _
					GetString(llngOffset + 1, 2) & " " & _
					CInt(GetHex(llngOffset + 3, 1)) & "." & _
					GetHex(llngOffset + 4, 1)

				.ExtractVersion			= FormatNumber(GetNumber(llngOffset + 5, 2) * .1, 1, True)
				.GeneralPurposeFlags	= GetNumber(llngOffset + 7, 2)
				.CompressionMethod		= GetNumber(llngOffset + 9, 2)
				.LastModifiedTime		= GetNumber(llngOffset + 11, 2)
				.LastModifiedDate		= GetNumber(llngOffset + 13, 2)
				.CRC32					= GetNumber(llngOffset + 15, 4)
				.CompressedSize			= GetNumber(llngOffset + 19, 4)
				.UncompressedSize		= GetNumber(llngOffset + 23, 4)
				.FileNameLength			= GetNumber(llngOffset + 27, 2)
				.ExtraFieldLength		= GetNumber(llngOffset + 29, 2)
				.FileName				= GetString(llngOffset + 31, .FileNameLength)
				.ExtraField				= GetString(llngOffset + 31 + .FileNameLength, .ExtraFieldLength)
				.StartByte				= llngOffSet + 1
				.EndByte				= llngOffSET + .FileNameLength + .ExtraFieldLength + .CompressedSize + 30
'				.BinaryData				= MidB(pbin_Zip, llngOffSET + .FileNameLength + .ExtraFieldLength + 30, .CompressedSize)
'				.LocalFileHeader		= GetString(llngOffset + 1, .FileNameLength + .ExtraFieldLength + 30)
				llngOffSet = .EndByte
				.IsOverall = (.Name = "" And .Path = "")
				.IsFolder = (.Name = "" And Not .Path = "")
			End With
			mlng_Files = mlng_Files + 1

		Loop While mobj_Files(mlng_Files - 1).EndByte < LenB(mbin_zip)
		
	End Sub
	
	Private Function GetHex(plngStart, plngLength)
		Dim llngIndex
		Dim lstrHex
		For llngIndex = 0 To plngLength - 1
			lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(mbin_zip, plngStart + llngIndex, 1))), 2)
		Next
		GetHex = lstrHex
	End Function
	Private Function GetString(plngStart, plngLength)
	
		Dim llngIndex
		Dim lstrString
		If LenB(mbin_zip) < (plngStart + (plngLength - 1)) Then Exit Function
		For llngIndex = 0 To plngLength - 1
			If AscB(MidB(mbin_zip, plngStart + llngIndex, 1)) = 0 Then
				'GetString = lstrString
				'Exit Function
				lstrString = lstrString & " "
			Else
				lstrString = lstrString & Chr(AscB(MidB(mbin_zip, plngStart + llngIndex, 1)))
			End If
		Next
		GetString = lstrString
	End Function
	Private Function GetNumber(plngStart, plngLength)
		If plngStart < 0 Then Exit Function
		Dim llngIndex
		Dim lstrHex
		For llngIndex = 0 To plngLength - 1
			lstrHex = Right("0" & Hex(AscB(MidB(mbin_zip, plngStart + llngIndex, 1))), 2) & lstrHex
		Next
		GetNumber = CDbl("&h" & lstrHex)
	End Function
	Function GetDate(plngStart)
		Dim llngDate
		llngDate = GetNumber(plngStart, 2)
		GetDate = DateSerial(1980 + (llngDate And &HFE00) \ &H200, (llngDate And &H1E0) \ &H20, llngDate And &H1F)
	End Function
	Function GetTime(plngStart)
		Dim llngDate
		llngDate = GetNumber(plngStart, 2)
		GetTime = TimeSerial((llngDate And &HF800) \ &H800, (llngDate And &H7E0) \ &H20, (llngDate And &H1F) * 2)
	End Function
'	TimeVal = Asc(Right$(OFS.FileTime, 1)) * 256& +_
'Asc(Left$(OFS.FileTime, 1))
'  S = (TimeVal And &H1F) * 2              ' Seconds
'  N = (TimeVal And &H7E0) \ &H20          ' Minutes
'  H = (TimeVal And &HF800) \ &H800        ' Hours
''
'' Parse Date value
''
'  DateVal = Asc(Right$(OFS.FileDate, 1)) * 256& +
'Asc(Left$(OFS.FileDate, 1))
'  D = (DateVal& And &H1F)             ' Days
'  M = (DateVal& And &H1E0) \ &H20     ' Months
'  Y = (DateVal& And &HFE00) \ &H200   ' Years'
'
'  GetFileCreateDateTime = DateSerial(1980 + Y, M, D) + TimeSerial(H,_
'N, S)
'
'	End Function

End Class
Class clsZipFile
	Public Signature				' 4
	Public ExtractVersion			' 2
	Public GeneralPurposeFlags		' 2
	Public CompressionMethod		' 2
	Public LastModifiedTime			' 2
	Public LastModifiedDate			' 2
	Public CRC32					' 4
	Public CompressedSize			' 4
	Public UncompressedSize			' 4
	Public FileNameLength			' 2
	Public ExtraFieldLength			' 2
	Public FileName					' see file name length
	Public ExtraField				' see extra field length
	Public StartByte				' 4
	Public EndByte					' 4
	Public BinaryData				' See CompressedSize
	Public LocalFileHeader
	
	Public IsFolder
	Public IsOverall
	
	Public Property Get Name
		Dim lstrPath
		lstrPath = Replace(FileName, "/", "\")
		If InStr(1, lstrPath, "\") = "0" Then
			Name = lstrPath
			Exit Property
		End If
		Name = Mid(lstrPath, InStrRev(lstrPath, "\") + 1)
	End Property
	Public Property Get Path
		Dim lstrPath
		lstrPath = Replace(FileName, "/", "\")
		If InStr(1, lstrPath, "\") = "0" Then
			Path = ""
			Exit Property
		End If
		Path = Mid(lstrPath, 1, InStrRev(lstrPath, "\"))
	End Property
	Public Property Get Packed
		Packed = CompressedSize
	End Property
	Public Property Get Ratio
		If UncompressedSize = 0 Then Exit Property
		If CompressedSize >= UncompressedSize Then
			Ratio = "0%"
		Else
			Ratio = FormatNumber(((1 - (CompressedSize / UncompressedSize)) * 100), 0, True, False, True) & "%"
		End If
	End Property
	Public Property Get Modified()
		Modified = CDate(GetDate(LastModifiedDate) & " " & GetTime(LastModifiedTime))
	End Property
	Private Function GetDate(plngDate)
		GetDate = DateSerial(1980 + (plngDate And &HFE00) \ &H200, _
			(plngDate And &H1E0) \ &H20, plngDate And &H1F)
	End Function
	Private Function GetTime(plngDate)
		GetTime = TimeSerial((plngDate And &HF800) \ &H800, _
			(plngDate And &H7E0) \ &H20, _
			(plngDate And &H1F) * 2)
	End Function
	Public Property Get Size()
		Size = UncompressedSize
	End Property
	Public Property Get BitMask()
		Dim llngNumber
		Dim lstrBits
		llngNumber = GeneralPurposeFlags
		Do
			If llngNumber Mod 2 = 1 Then lstrBits = "1" & lstrBits Else lstrBits = "0" & lstrBits
			llngNumber = llngNumber \ 2
		Loop Until llngNumber = 0
		lstrBits = Right("0000000000000000" & lstrBits, 16)
		For llngNumber = 0 To 3
			lstrReturn = lstrReturn & Mid(lstrBits, (llngNumber * 4) + 1, 4) & "."
		Next
		BitMask = Left(lstrReturn, 19)
	End Property
	Property Get CompressionMethodString()
		Select Case CompressionMethod
			Case 0
				CompressionMethodString = "The file is stored (no compression)"
			Case 1
				CompressionMethodString = "The file is Shrunk"
			Case 2
				CompressionMethodString = "The file is Reduced with compression factor 1"
			Case 3
				CompressionMethodString = "The file is Reduced with compression factor 2"
			Case 4
				CompressionMethodString = "The file is Reduced with compression factor 3"
			Case 5
				CompressionMethodString = "The file is Reduced with compression factor 4"
			Case 6
				CompressionMethodString = "The file is Imploded"
			Case 7
				CompressionMethodString = "Reserved for Tokenizing compression algorithm"
			Case 8
				CompressionMethodString = "The file is Deflated"
			Case 9
				CompressionMethodString = "Reserved for enhanced Deflating"
			Case 10
				CompressionMethodString = "PKWARE Date Compression Library Imploding"
			Case Else
				CompressionMethodString = "Unhandled Copression type: " & CompressionMethod
		End Select
	End Property
End Class
%>