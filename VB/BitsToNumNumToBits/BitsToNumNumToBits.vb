'**************************************
' Name: BitsToNum/NumToBits
' Description:I was reading some binary 
'     files and needed some routines to conver
'     t numbers back and forth into binary num
'     bers. In this case, I set the procedures
'     to work with binary arrays.
' By: Lewis Moten
'Assumes:Assumes you are familliar with 
'     creating arrays. some understanding of b
'     inary conversion would prove helpful.
'This code is copyrighted and has limite
'     d warranties.
'**************************************

Function BitsToNum(ByRef pbitArray() As Boolean) As Long
    Dim lnMaxBits As Long
    Dim lnIndex As Long
    lnMaxBits = UBound(pbitArray)
    BitsToNum = 0
    For lnIndex = 0 To lnMaxBits
        If pbitArray(lnIndex) Then
            BitsToNum = BitsToNum + (2 ^ (lnMaxBits - lnIndex))
        End If
    Next
End Function

Private Sub NumToBits(ByVal pNumber As Long, ByRef pbitArray() As Boolean)
    Dim lnIndex As Long
    lnIndex = UBound(pbitArray)
    Do
        pbitArray(lnIndex) = pNumber Mod 2 = 1
        pNumber = pNumber \ 2
        lnIndex = lnIndex - 1
    Loop Until pNumber = 0
End Sub