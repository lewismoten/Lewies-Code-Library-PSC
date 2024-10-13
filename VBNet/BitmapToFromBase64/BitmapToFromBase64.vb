'**************************************
' Name: BitmapTo/FromBase64
' Description:Allows you to convert bitmap images to base64 and back for storage and retrieval in XML files.
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
'This code is copyrighted and has limited warranties.
'**************************************
    
Public Function BitmapToBase64(ByVal image As System.Drawing.Bitmap) As String
	Dim base64 As String
	Dim memory As New System.IO.MemoryStream()
	image.Save(memory, Imaging.ImageFormat.Bmp)
	base64 = System.Convert.ToBase64String(memory.ToArray)
	memory.Close()
	memory = Nothing
	Return base64
End Function

Public Function BitmapFromBase64(ByVal base64 As String) As System.Drawing.Bitmap
	Dim oBitmap As System.Drawing.Bitmap
	Dim memory As New System.IO.MemoryStream(Convert.FromBase64String(base64))
	oBitmap = New System.Drawing.Bitmap(memory)
	memory.Close()
	memory = Nothing
	Return oBitmap
End Function