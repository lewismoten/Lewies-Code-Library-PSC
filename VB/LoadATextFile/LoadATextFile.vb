'**************************************
' Name: Load a Text File
' Description:Very Simple (beginners code) - reads and returns the contents of a text file.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************
MsgBox TextFile("C:\autoexec.bat")

Function TextFile(ByRef pStrFileName As String)
    Dim llngFileNum As Long
    Dim llngFileLen As Long
    On Error Resume Next
    ' Get the size of the file
    ' If the file does not exist, an error w
    '     ill occur
    llngFileLen = 0
    llngFileLen = FileLen(pStrFileName)
    If llngFileLen = 0 Then Exit Function
    On Error GoTo 0
    ' Read the file into memory
    llngFileNum = FreeFile
    Open pStrFileName For Input As #llngFileNum
    TextFile = Input(llngFileLen, #llngFileNum)
    Close #lngFileNum
End Function
		
