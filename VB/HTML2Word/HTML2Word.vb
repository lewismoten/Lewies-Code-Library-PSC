'**************************************
' Name: HTML2Word dll
' Description:Converts a web page to a word document.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Option Explicit
' Make sure the following references are in your project before compiling.
' Also - the server that you install this on must also have these products
' installed.
'
' If your server spits back an error, itis probably because the Word object
' is nost installed.
'Microsoft Word 9.0 Object Library
'C:\Program Files\Microsoft Office\Office\MSWORD9.OLB
'Microsoft Active Server Pages Object Library
'C:\WINNT\System32\inetsvr\asp.dll
'Microsoft ActiveX Data Objects 2.6 Library
'C:\Program Files\Common Files\System\ADO\msado15.dll
'Microsoft Internet Controls
'C:\WINNT\System32\SHDOCVW.DLL
Private ASP As ASPTypeLibrary.ScriptingContext
Private Response As ASPTypeLibrary.Response
Private Session As ASPTypeLibrary.Session
Private Server As ASPTypeLibrary.Server
Private WithEvents IE As SHDocVw.InternetExplorer
Private Word As Word.Document
Private Stream As ADODB.Stream
Private mblnDone


Public Sub OnStartPage(ByRef ASPLink As ASPTypeLibrary.ScriptingContext)
    Set ASP = ASPLink
    Set Response = ASPLink.Response
    Set Session = ASPLink.Session
    Set Server = ASPLink.Server
    Set IE = New SHDocVw.InternetExplorer
    Set Word = New Word.Document
    Set Stream = New ADODB.Stream
    Response.Clear
End Sub


Private Sub Cleanup()
    Set IE = Nothing
    Set Word = Nothing
    Set Response = Nothing
    Set Session = Nothing
    Set Server = Nothing
    Set ASP = Nothing
    Set Stream = Nothing
End Sub


Public Sub Download(ByRef pstrURL As Variant)
    Dim lstrPath As String
    Dim lstrFileName As String
    Dim ldblStart As Double
    mblnDone = False
    ldblStart = Timer
    Call IE.Navigate2(pstrURL)

    While IE.Busy And Not mblnDone
		DoEvents
		If (Timer - ldblStart) > Server.ScriptTimeout Then
			Call Cleanup
			Err.Raise vbObjectError + 1, "HTML2Word.dll", "Connect Timeout - Busy"
		End If
	Wend


	While Not (IE.Document.ReadyState = "complete" Or mblnDone)
        DoEvents
        If (Timer - ldblStart) > Server.ScriptTimeout Then
            Call Cleanup
            Err.Raise vbObjectError + 2, "HTML2Word.dll", "Connect Timeout - Not Complete"
        End If
    Wend
    Call IE.Document.Body.createTextRange.execCommand("Copy")


    DoEvents
    lstrFileName = Session.SessionID & ".doc"
    lstrPath = App.Path & "\~" & Hex(Timer) & "_" & lstrFileName


	DoEvents
	On Error Resume Next
	Word.Content.Paste

	If Err Then
		Call Cleanup
		Dim lstrMsg
		lstrMsg = Err.Description
		On Error GoTo 0
		Err.Raise vbObjectError + 3, "HTML2Word.dll", "Can not paste - " & lstrMsg
	End If
	On Error GoTo 0
	Word.SaveAs lstrPath
	Word.Close
	Response.ContentType = "application/octet-stream"
	Response.AddHeader "content-disposition", "attatchment; filename=" & lstrFileName
	Stream.Open
	Stream.LoadFromFile lstrPath
	Response.BinaryWrite Stream.ReadText
	Stream.Close
	Response.Flush
	Response.End
	FileSystem.Kill lstrPath
End Sub

Public Sub OnEndPage()
    Call Cleanup
End Sub

Private Sub IE_StatusTextChange(ByVal Text As String)
    If Text = "Done" Then mblnDone = True
    DoEvents
End Sub