'**************************************
' Name: Average Color In Image
' Description:Returns the average color 
'     within a picture box. I use this when cr
    '     eating small maps of tile-based rpg game
    '     s. I get the avg color of each block, an
    '     d place the color on the map.
' By: Lewis Moten
'This code is copyrighted and has limite
'     d warranties.
'**************************************



Function AverageRGB(ByRef P As PictureBox) As Long
    Dim Count As Long
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    Dim Hexed As String
    Dim X As Long
    Dim Y As Long
    Count = 0


    For X = 0 To P.Width Step P.Width \ 32


        For Y = 0 To P.Height Step P.Height \ 32
            Hexed = Right("00000" & Hex(P.Point(X, Y)), 6)
            Red = Red + CLng("&h" & Right(Hexed, 2))
            Green = Green + CLng("&h" & Mid(Hexed, 3, 2))
            Blue = Blue + CLng("&h" & Left(Hexed, 2))
            Count = Count + 1
        Next
    Next
    AverageRGB = RGB(Red \ Count, Green \ Count, Blue \ Count)
End Function