'**************************************
'Windows API/Global Declarations for :PicMatch2
'**************************************
' API to copy image to memory
'
'Private Declare Function GetDIBits Lib "gdi32" _
'    ( _
'    ByVal aHDC As Long, _
'    ByVal hBitmap As Long, _
'    ByVal nStartScan As Long, _
'    ByVal nNumScans As Long, _
'    lpBits As Any, _
'    lpBI As BitmapInfo, _
'    ByVal wUsage As Long _
'    ) As Long
'**************************************
' Name: PicMatch2
' Description:Compares each pixels colors in one picture box to another. Determines
' if pictures are an exact match. Now uses API to copy pictures to memory. Looking
' for a quicker way to compare arrays. Copy this code into a module and call PicMatch.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Private Type BitmapInfoHeader
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type

Private Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Reserved As Byte
End Type

Private Type BitmapInfo
    Header As BitmapInfoHeader
    Colors(0 To 15) As RGBQuad
End Type

Private Const PIXEL As Integer = 3
Private Const DIB_RGB_COLORS As Long = 0
Private Const PALVERSION As Long = &H300

' API to copy image to memory
Private Declare Function GetDIBits Lib "gdi32" _
    ( _
    ByVal aHDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    lpBits As Any, _
    lpBI As BitmapInfo, _
    ByVal wUsage As Long _
    ) As Long

Public Function PicMatch(ByRef Pic1 As PictureBox, ByRef pic2 As PictureBox) As Boolean
    Dim p1() As Byte
    Dim p2() As Byte
    ' copy both picturs to memory
    p1 = BMPArray(Pic1)
    p2 = BMPArray(pic2)
    ' compare arrays of image data
    PicMatch = ArraysMatch(p1, p2)
End Function

Private Function BMPArray(ByRef Pic As PictureBox) As Byte()
    Dim LocalBitmapInfo As BitmapInfo
    Dim LocalFileHeader As BitmapInfoHeader
    Dim Bits() As Byte
    Dim BufferSize As Long
    Dim ReturnValue As Long
    Const BitsPixel As Long = 4
    ' Setup buffer to hold data
    BufferSize = ((Pic.ScaleWidth / 2 + 3) And &HFFFC) * Pic.ScaleHeight
    ReDim Bits(0 To BufferSize - 1)
    ' Setup header properties


    With LocalBitmapInfo.Header
        .Size = 40
        .Width = Pic.ScaleWidth
        .Height = Pic.ScaleHeight
        .Planes = 1
        .BitCount = BitsPixel
        .Compression = 0
        .ClrUsed = 0
        .ClrImportant = 0
        .SizeImage = BufferSize
    End With
    ' Copy bitmap to memory
    ReturnValue = GetDIBits _
    ( _
    Pic.hDC, _
    Pic.Image, _
    0, _
    LocalBitmapInfo.Header.Height, _
    Bits(0), _
    LocalBitmapInfo, _
    DIB_RGB_COLORS _
    )
    ' Return the image portion
    BMPArray = Bits()
End Function

Private Function ArraysMatch(ByRef Array1 As Variant, ByRef Array2 As Variant) As Boolean
    Dim Low As Long
    Dim High As Long
    Dim Index As Long
    ' Get lower and upper bounds
    Low = LBound(Array1)
    High = UBound(Array1)
    ' make sure second array matches the lower and upper bounds
    If Not (LBound(Array2) = Low And UBound(Array2) = High) Then
        ArraysMatch = False
        Exit Function
    End If
    ' loop through each index
    For Index = Low To High
        ' if no matches - return negative results
        If Not Array1(Index) = Array2(Index) Then
            ArraysMatch = False
            Exit Function
        End If
    Next
    ' If we got this far, then we have been successful.
    ArraysMatch = True
End Function