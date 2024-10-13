'**************************************
' Name: Simple Log Routine
' Description:Simple routine that logs a
'     sting to a specified log file. also supp
'     orts word-wrapping on long lines (more t
'     hen 80 characters) and inserts an empty 
'     line before and after the information so
'     it is easier to read. This code is mainl
'     y here for my own personal reference.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Public Sub LogInfo(ByRef pStrInfo As String)
    Dim lLngFileNum As Long
    Dim lLngLength As Long
    Dim lStrLogFile As String
    Const lLngMaxChars = 80
    lLngFileNum = FreeFile
    lLngLength = Len(pStrInfo)
    lLngPosition = 1
    lStrLogFile = txtLogFile.Text ' Change this line to the appropriate log file location
    On Error Resume Next
    Open lStrLogFile For Append Access Write As #lLngFileNum
    If lLngLength > lLngMaxChars Then Print #lLngFileNum, ""
    Do
        Print #lLngFileNum, Mid(pStrInfo, lLngPosition, lLngMaxChars)
        lLngPosition = lLngPosition + lLngMaxChars
    Loop While lLngPosition <= lLngLength
    If lLngLength > lLngMaxChars Then Print #lLngFileNum, ""
    Close #lLngFileNum
End Sub