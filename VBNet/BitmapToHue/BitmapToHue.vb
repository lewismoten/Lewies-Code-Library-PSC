	'**************************************
    ' Name: Bitmap to Hue
    ' Description:Supply a bitmap and color and the bitmap is magically transformed to show colors from black to your color. Works with GIF images. Great for greyscale backgrounds that need to keep image but only change shades of color.
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
    
    	Private Function BitmapToHue(ByVal oBitmap As Drawing.Bitmap, ByVal oHue As Color) As Drawing.Bitmap
    		Dim iIndex As Integer
    		Dim iMaxIndex As Integer = oBitmap.Palette.Entries.Length - 1
    		Dim oEntry As Drawing.Color
    		Dim oPalette As Drawing.Imaging.ColorPalette
    		Dim duPercent As Double
    		Dim duAverage As Double
    		Dim iR As Byte	' Red Intensity
    		Dim iG As Byte	' Green Intensity
    		Dim iB As Byte	' Blue Intensity
    		oPalette = oBitmap.Palette
    		For iIndex = 0 To iMaxIndex
    			oEntry = oPalette.Entries(iIndex)
    			duAverage = (CInt(oEntry.R) + CInt(oEntry.R) + CInt(oEntry.G)) / 3
    			duPercent = duAverage / 255
    			iR = CType(oHue.R * duPercent, Byte)
    			iG = CType(oHue.G * duPercent, Byte)
    			iB = CType(oHue.B * duPercent, Byte)
    			oPalette.Entries(iIndex) = Color.FromArgb(255, iR, iG, iB)
    		Next
    		' Setup new palette
    		oBitmap.Palette = oPalette
    		' Return the new image (applies palette with "new")
    		Return New Bitmap(oBitmap)
    	End Function

		
