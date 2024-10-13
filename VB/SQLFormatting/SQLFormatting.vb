'**************************************
' Name: SQL Formatting
' Description:Format variables for SQL Syntax.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************
Public Function SQLNumber(ByVal pvarNumber As Variant, Optional ByRef pblnForceZero As Boolean = False) As String
    Dim lstrNum As String
    lstrNum = CStr(pvarNumber & "")
    If IsNumeric(lstrNum) Then
        SQLNumber = lstrNum
    Else
        If pblnForceZero Then
            SQLNumber = CStr(ParseNumber(lstrNum))
        Else
            SQLNumber = "NULL"
        End If
    End If
End Function

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
    lblnNegative = InStr(1, pvarNumber, "-") = 1
    ' determine if negative (parens)
    If Not lblnNegative Then
        lblnNegative = InStr(1, pstrText, "(") = 1 And InStr(1, pstrText, ")") = Len(pstrText)
    End If
    ' Count characters
    llngMaxIndex = Len(pstrText)
    ' Loop through each character
    For llngIndex = 1 To llngMaxIndex
        ' if character is a digit or decimal
        If Not InStr(1, lstrNum, Mid(pvarNumber, llngIndex, 1)) = 0 Then
            ' append character to "number"
            lstrNumber = lstrNumber & Mid(pvarNumber, llngIndex, 1)
        End If
    Next ' llngIndex
    If lblnNegative Then lstrNumber = "-" & lstrNumber
    ParseNumber = CDbl(lstrNumber)
End Function

Public Function SQLBoolean(ByRef pvarBoolean As Variant) As String
    Dim lblnResult As Boolean
    If IsNull(pvarBoolean) Then
        SQLBoolean = "0"
        Exit Function
    End If
    On Error Resume Next
    lblnResult = CBool(pvarBoolean)
    If Err Then
        SQLBoolean = "0"
        Exit Function
    End If
    If lblnResult Then
        SQLBoolean = "-1"
    Else
        SQLBoolean = "0"
    End If
End Function

Public Function SQLDate(ByRef pvarDate As Variant) As String
    If IsDate(pvarDate) Then
        SQLDate = "'" & pvarDate & "'"
    Else
        SQLDate = "NULL"
    End If
End Function

Public Function SQLString(ByRef pvarString As Variant, Optional ByRef plngMaxLength As Long = -1) As String
    Dim lstrString As String
    If IsNull(pvarString) Then
        SQLString = "NULL"
        Exit Function
    End If
    lstrString = Trim(CStr(pvarString))
    If Len(lstrString) = 0 Then
        SQLString = "NULL"
        Exit Function
    End If
    If plngMaxLength > 0 Then
        lstrString = Left(lstrString, plngMaxLength)
    End If
    lstrString = Replace(lstrString, "'", "''")
    lstrString = "'" & lstrString & "'"
    SQLString = lstrString
End Function