'**************************************
' Name: ParseNumber
' Description:Parses a number from a string. Handles parans for negative numbers.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************
Public Function ParseNumber(ByVal pstrText As String) As Double
    Dim lblnNegative As Boolean
    Dim lstrNumber As String
    Dim llngIndex As Long
    Dim llngMaxIndex As Long
    Const lstrNum = "0123456789."
    ' Initialize number
    lstrNumber = "0"
    ' remove currency
    pstrText = Replace(pstrText, "$", "")
    ' determine if negative
    lblnNegative = InStr(1, pstrText, "-") = 1
    ' determine if negative (parens)
    If Not lblnNegative Then
        lblnNegative = InStr(1, pstrText, "(") = 1 And InStr(1, pstrText, ")") = Len(pstrText)
    End If
    ' Count characters
    llngMaxIndex = Len(pstrText)
    ' Loop through each character
    For llngIndex = 1 To llngMaxIndex
        ' if character is a digit or decimal
        If Not InStr(1, lstrNum, Mid(pstrText, llngIndex, 1)) = 0 Then
            ' append character to "number"
            lstrNumber = lstrNumber & Mid(pstrText, llngIndex, 1)
        End If
    Next ' llngIndex
    If lblnNegative Then lstrNumber = "-" & lstrNumber
    ParseNumber = CDbl(lstrNumber)
End Function