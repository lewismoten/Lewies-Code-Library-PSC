'**************************************
' Name: Dynamic Intrinsic Controls
' Description:This demo shows you how to
'     dynamically create a FileListBox control
'     via vb Code rather then at design time. 
'     The reason I choose to create this was t
'     hat I needed to get a list of files with
'     in a directory via. a DLL. The file syst
'     em object causes too much overhead and t
'     his little trick makes the code run much
'     faster. For testing purposes, just creat
'     e a new EXE and throw this onto Form1. F
'     or more info, See the MS Knowledge Base:
'     Q190670
' By: Lewis Moten
'This code is copyrighted and has limite
'     d warranties.
'**************************************
Private Sub Form_Load()
    Dim File1 As Object
    Dim Max As Long
    Dim i As Long
    Dim msg As String
    Set File1 = Controls.Add("VB.FileListBox", "File1", Form1)
    File1.Path = "C:\"
    Max = File1.ListCount
    For i = 0 To Max - 1
        msg = msg & File1.List(i) & vbCrLf
    Next
    MsgBox msg
End Sub
		
