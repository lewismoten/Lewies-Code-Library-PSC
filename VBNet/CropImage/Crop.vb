' Name: Crop Image
' Description:Crops an image passed to 
'     the function. Returns resized image as a
'     new bitmap.
' By: Lewis Moten
'
'This code is copyrighted and has limited warranties.

Private Function Crop(ByVal Source As Bitmap, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32) As Bitmap
    Dim Cropped As New Bitmap(width, height)
    Dim g As Graphics = Graphics.FromImage(Cropped)
    g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
    g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
    Dim rect As New Rectangle(0, 0, width, height)
    g.DrawImage(Source, rect, x, y, width, height, GraphicsUnit.Pixel)
    Return Cropped
End Function