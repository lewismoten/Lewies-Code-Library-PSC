'**************************************
' Name: MakePath
' Description:Creates folders and subfolders to the specified path.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Option Explicit
' Make sure path is ok

If Not PathOK(Path) Then
    MsgBox "The task could not be completed because the destination is invalid."
    Exit Sub
End If

Public Function PathOK(ByVal pstrPath As String) As Boolean
    Dim llngAnswer As VbMsgBoxResult
    Dim llngTemp As Long
    On Error Resume Next
    ' Strip trailing slash
    If Right(pstrPath, 1) = "\" Then Path = Left(pstrPath, Len(pstrPath) - 1)
    ' Causes error if invalid ...
    llngTemp = FileSystem.FileLen(pstrPath)
    ' if directory does not exist


    If Err Then
        On Error GoTo 0
        ' ask if we should create a new directory
        llngAnswer = MsgBox("The path does not exist. Should the path be created?", vbOKCancel, "Path Not Found")
        If llngAnswer = vbOK Then
            PathOK = MakePath(pstrPath)
        Else
            PathOK = False
        End If
        ' else directory exists ...
    Else
        PathOK = True
    End If
End Function

Public Function MakePath(ByVal pstrPath As String) As Boolean
    On Error Resume Next
    Dim lstrParent As String
    ' Remove possible trailing slash
    If Right(pstrPath, 1) = "\" Then pstrPath = Left(pstrPath, Len(pstrPath) - 1)
    ' Has parent directory?

    If InStrRev(pstrPath, "\") = 0 Then
        MakePath = False
        Exit Function
    End If
    ' Attempt to create folder
    FileSystem.MkDir pstrPath & "\"

    If Err Then
        ' parse parent
        lstrParent = Mid(pstrPath, 1, InStrRev(pstrPath, "\") - 1)
        ' attempt to make parent directory
        If MakePath(lstrParent) Then MakePath = MakePath(pstrPath) Else MakePath = False
    Else
        MakePath = True
        Exit Function
    End If
End Function