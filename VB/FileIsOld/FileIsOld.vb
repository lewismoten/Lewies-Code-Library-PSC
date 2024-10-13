'**************************************
' Name: Determine if File Is Old
' Description:Determines if a file is ol
'     d. I use this when I loop through the fi
'     les in a "temp" directory to determine i
'     f I should delete old files on a website
'     . Take note - the function looks at the 
'     last modified date rather then the date 
'     created.
' By: Lewis Moten
'This code is copyrighted and has limite
'     d warranties.
'**************************************
Private Sub Form_Load()
    MsgBox IIf(FileIsOld("C:\AutoExec.bat"), "The file is old", "The file is new")
End Sub

Function FileIsOld(ByRef pStrFilePath As String) As Boolean
    Dim llngMinutesOld As Long
    Dim ldtmLastModified As Date
    Dim llngFileAttr As VbFileAttribute
    Const llngMinutesOldAfter As Long = 10
    On Error Resume Next
    llngFileAttr = FileSystem.GetAttr(pStrFilePath)


    If Err Then
        MsgBox "File does not exist."
        Exit Function ' file doesn't exist
    End If
    On Error GoTo 0
    If Len(FileSystem.Dir(pStrFilePath, llngFileAttr)) = 0 Then Exit Function
    ldtmLastModified = FileSystem.FileDateTime(pStrFilePath)
    llngMinutesOld = DateDiff("n", ldtmLastModified, Now())
    FileIsOld = llngMinutesOld > pLngMinutesOldAfter
End Function