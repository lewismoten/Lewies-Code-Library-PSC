'**************************************
' Name: PicMatch
' Description:Compares each pixels color
'     s in one picture box to another. Determi
'     nes if pictures are an exact match. Look
'     ing for a quicker way to perform functio
'     nality if possible.
' By: Lewis Moten
'Side Effects:None
'This code is copyrighted and has limited warranties.
'**************************************
Function PicMatch(ByRef P1 As PictureBox, ByRef P2 As PictureBox) As Boolean
    For X = 0 To P1.Width
        For Y = 0 To P1.Height
            If Not P1.Point(X, Y) = P2.Point(X, Y) Then
                PicMatch = False
                Exit Function
            End If
        Next
    Next
    PicMatch = True
End Function