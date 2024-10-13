'**************************************
' Name: Convert 16 bit colors to 24 bit
' Description:I was working with some fi
'     les from a game and the colors were in 1
'     6 bit. I had to convert them to 24 bit j
'     ust so that I could view them. These fun
'     ctions do just that.
' By: Lewis Moten
' Inputs:pnColor16 - Numerical represent
'     ation of 16 bit color.
' Returns:Color24 - Numerical representa
'     tion of 24 bit color.
'This code is copyrighted and has limite
'     d warranties.
'**************************************

nColor16 = (lnByte1 * (2 ^ 8)) + lnByte2
nColor24 = Color16To24(nColor16)

Function Color16To24(ByRef pnColor16 As Long) As Long
    
    Dim lnRed As Long
    Dim lnGreen As Long
    Dim lnBlue As Long
    
    lnRed = ((RShift(pnColor16, 10) And &H1F) * &HFF \ &H1F)
    lnGreen = LShift(((RShift(pnColor16, 5) And &H1F) * &HFF \ &H1F), 8)
    lnBlue = LShift(((pnColor16 And &H1F) * &HFF \ &H1F), 16)
    Color16To24 = (lnRed Or lnGreen Or lnBlue)
End Function

Private Function LShift(ByRef pnValue As Long, ByRef pnShift As Long)
    ' Equivilant to C's Bitwise << operator
    LShift = pnValue * (2 ^ pnShift)
End Function

Private Function RShift(ByRef pnValue As Long, ByRef pnShift As Long)
    ' Equivilant to C's Bitwise >> operator
    RShift = CLng(pnValue \ (2 ^ pnShift))
End Function