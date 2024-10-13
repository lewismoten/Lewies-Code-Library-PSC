'**************************************
' Name: Left Shift and Right Shift
' Description:Allows you to "Shift" bits
'     around. These functions mimic the bitwis
'     e opperators << and >> in C.
' By: Lewis Moten
' Inputs:
'	pnValue - the byte, or number to manipulate.
'	pnShift - position to shift bits to. (wraps)
' Returns:Shifted Number
'Assumes:Assumes you have a little under
'     standing of Bitwise manipulation.
'This code is copyrighted and has limite
'     d warranties.
'**************************************

Private Function LShift(ByRef pnValue As Long, ByRef pnShift As Long)
    ' Equivilant to C's Bitwise << operator
    LShift = pnValue * (2 ^ pnShift)
End Function

Private Function RShift(ByRef pnValue As Long, ByRef pnShift As Long)
    ' Equivilant to C's Bitwise >> operator
    RShift = CLng(pnValue \ (2 ^ pnShift))
End Function