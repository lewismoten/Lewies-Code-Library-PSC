<%
' ------------------------------------------------------------------------------
' Container of Field Properties
Class clsField
	Public FileName
	Public ContentType
	Public Value
	Public FieldName
	Public Length
	Public BinaryData
End Class
' ------------------------------------------------------------------------------
Class clsUpload
' ------------------------------------------------------------------------------
	Private nFieldCount
	Private oFields()
	
' ------------------------------------------------------------------------------
	Public Property Get Count()
		Count = nFieldCount
	End Property
' ------------------------------------------------------------------------------
	Public Default Property Get Field(ByRef asFieldName)
		Dim lnLength
		Dim lnIndex
		
		lnLength = UBound(oFields)
		
		If IsNumeric(asFieldName) Then
			If lnLength >= asFieldName And asFieldName > -1 Then
				Set Field = oFields(asFieldName)
			Else
				Set Field = New clsField
			End If
		Else
			For lnIndex = 0 To lnLength
				If LCase(oFields(lnIndex).FieldName) = LCase(asFieldName) Then
					Set Field = oFields(lnIndex)
					Exit Property
				End If
			Next
			Set Field = New clsField
		End If
	End Property
' ------------------------------------------------------------------------------
	Public Function Exists(ByRef avKeyIndex)
		Exists = Not IndexOf(avKeyIndex) = -1
	End Function
' ------------------------------------------------------------------------------
	Public Property Get ValueOf(ByRef avKeyIndex)
		Dim lnIndex
		lnIndex = IndexOf(avKeyIndex)
		if lnIndex = -1 Then Exit Property
		ValueOf = oFields(lnIndex).Value
	End Property
' ------------------------------------------------------------------------------
	Public Property Get FileNameOf(ByRef avKeyIndex)
		Dim lnIndex
		lnIndex = IndexOf(avKeyIndex)
		if lnIndex = -1 Then Exit Property
		FileNameOf = oFields(lnIndex).FileName
	End Property
' ------------------------------------------------------------------------------
	Public Property Get LengthOf(ByRef avKeyIndex)
		Dim lnIndex
		lnIndex = IndexOf(avKeyIndex)
		if lnIndex = -1 Then Exit Property
		LengthOf = oFields(lnIndex).LengthOf
	End Property
' ------------------------------------------------------------------------------
	Public Property Get BinaryDataOf(ByRef avKeyIndex)
		Dim lnIndex
		lnIndex = IndexOf(avKeyIndex)
		if lnIndex = -1 Then Exit Property
		BinaryDataOf = oFields(lnIndex).BinaryData
	End Property
' ------------------------------------------------------------------------------
	Private Function IndexOf(ByVal avKeyIndex)
		Dim lnIndex
		
		If avKeyIndex = "" Then
			IndexOf = -1
		ElseIf IsNumeric(avKeyIndex) Then
			avKeyIndex = CLng(avKeyIndex)
			If nFieldCount > avKeyIndex And avKeyIndex > -1 Then
				IndexOf = avKeyIndex
			Else
				IndexOf = -1
			End If
		Else
			For lnIndex = 0 To nFieldCount - 1
				If LCase(oFields(lnIndex).FieldName) = LCase(avKeyIndex) Then
					IndexOf = lnIndex
					Exit Function
				End If
			Next
			IndexOf = -1
		End If
	End Function
' ------------------------------------------------------------------------------
	Public Property Get ContentTypeOf(ByRef avKeyIndex)
		Dim lnIndex
		lnIndex = IndexOf(avKeyIndex)
		if lnIndex = -1 Then Exit Property
		ContentTypeOf = oFields(lnIndex).ContentType
	End Property
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		Dim lnIndex
		For lnIndex = 0 To nFieldCount - 1
			Set oFields(0) = Nothing
		Next
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		Dim lnBytes				' Bytes received from the client
		Dim lnByteCount			' Number of bytes received
		Dim lnStartPosition		' Position at which content begins
		Dim lnEndPosition		' Position at which content ends
		
		Dim loDic				' Contains properties of each
								' specific field
								' Local dictionary object(s) 
								' to be appended to class-scope
								' dictioary object.
								
		Dim lnBoundaryBytes		' Bytes contained within the current boundary
		Dim lnBoundaryStart		' Position at wich the current boundary begins
								' within the lnBytes binary data.
		Dim lnBoundaryEnd		' Position at wich the current boundary ends
								' within the lnBytes binary data.
		Dim lnDispositionPosition
		
		Dim lsFieldName			' Name of the current field being parsed from
								' Binary Data
		Dim lsFileName			' Name of the file within the current boundary
		Dim lnFileNamePosition	' Location of file name within current boundary
		Dim loField				' clsField Object
		Dim lsValue				' Value of the current field
		Dim lsContentType		' ContentType of the binary file (MIME Type)
		
		' Initialize Fields
		nFieldCount = 0
		ReDim oFields(-1)
		
		' Read the bytes (binary data) into memory	
		lnByteCount = Request.TotalBytes
		lnBytes = Request.BinaryRead(lnByteCount)
		
		'Get the lnBoundaryBytes
		lnStartPosition = 1
		lnEndPosition = InstrB(lnStartPosition, lnBytes, CStrB(vbCr))
		
		If lnEndPosition >= lnStartPosition Then
			lnBoundaryBytes = MidB(lnBytes, lnStartPosition, lnEndPosition - lnStartPosition)
		End If
		
		lnBoundaryStart = InstrB(1, lnBytes, lnBoundaryBytes)
		
		
		' Loop until the BoundaryBytes begin with "--"
		Do Until (lnBoundaryStart = InstrB(lnBytes, lnBoundaryBytes & CStrB("--")))
		
			' All data within this boundary is stored within a local dictionary
			' to be appended to the class-scope dictionary.
			
			ReDim Preserve oFields(nFieldCount)
			nFieldCount = nFieldCount + 1
			
			Set loField = New clsField

			lnDispositionPosition = InstrB(lnBoundaryStart, lnBytes, CStrB("Content-Disposition"))
			
			' Get an object name
			lnStartPosition = InstrB(lnDispositionPosition, lnBytes, CStrB("name=")) + 6
			lnEndPosition = InstrB(lnStartPosition, lnBytes, CStrB(""""))
			lsFieldName = CStrU(MidB(lnBytes, lnStartPosition, lnEndPosition - lnStartPosition))
			loField.FieldName = lsFieldName
			
			' Get the location fo the file name.
			lnFileNamePosition = InstrB(lnBoundaryStart, lnBytes, CStrB("filename="))
			lnBoundaryEnd = InstrB(lnEndPosition, lnBytes, lnBoundaryBytes)
			
			'Test if object is a file
			If Not lnFileNamePosition = 0 And lnFileNamePosition < lnBoundaryEnd Then
			
				' Parse Filename
				lnStartPosition = lnFileNamePosition + 10
				lnEndPosition =  InstrB(lnStartPosition, lnBytes, CStrB(""""))
				lsFileName = CStrU(MidB(lnBytes,lnStartPosition,lnEndPosition-lnStartPosition))
				loField.FileName = lsFileName				
				
				' Parse Content-Type
				lnStartPosition = InstrB(lnEndPosition,lnBytes,CStrB("Content-Type:")) + 14
				lnEndPosition = InstrB(lnStartPosition,lnBytes,CStrB(vbCr))
				lsContentType = CStrU(MidB(lnBytes,lnStartPosition,lnEndPosition-lnStartPosition))
				loField.ContentType = lsContentType

				' Parse Content
				lnStartPosition = lnEndPosition + 4
				lnEndPosition = InstrB(lnStartPosition,lnBytes,lnBoundaryBytes)-2
				lsValue = MidB(lnBytes,lnStartPosition,lnEndPosition-lnStartPosition)
				loField.BinaryData = lsValue & CStrB(vbNull)
				loField.Length = LenB(lsValue)
			Else

				' Parse Content
				lnStartPosition = InstrB(lnDispositionPosition, lnBytes, CStrB(vbCr)) + 4
				lnEndPosition = InstrB(lnStartPosition, lnBytes, lnBoundaryBytes) - 2
				lsValue = CStrU(MidB(lnBytes,lnStartPosition,lnEndPosition-lnStartPosition))
				loField.Value = lsValue
				loField.Length = Len(lsValue)
			End If

			Set oFields(UBound(oFields)) = loField

			'Loop to next object
			lnBoundaryStart = InstrB(lnBoundaryStart + LenB(lnBoundaryBytes), lnBytes, lnBoundaryBytes)
			
			Set loField = Nothing
			
		Loop

	End Sub
' ------------------------------------------------------------------------------
	Private Function CStrU(ByRef psByteString)
		Dim lnLength
		Dim lnPosition
		lnLength = LenB(psByteString)
		For lnPosition = 1 To lnLength
			CStrU = CStrU & Chr(AscB(MidB(psByteString, lnPosition, 1)))
		Next
	End Function
' ------------------------------------------------------------------------------
	Private Function CStrB(ByRef psUnicodeString)
		Dim lnLength
		Dim lnPosition
		lnLength = Len(psUnicodeString)
		For lnPosition = 1 To lnLength
			CStrB = CStrB & ChrB(AscB(Mid(psUnicodeString, lnPosition, 1)))
		Next
	End Function
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>