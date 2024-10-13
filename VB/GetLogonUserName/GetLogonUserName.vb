'**************************************
' Name: Get Logon User Name
' Description:Reads the logon name of th
'     e current user. Tested with NT. Win95 ap
'     pears to work as long as a user logs int
'     o the computer.
' By: Lewis Moten
'This code is copyrighted and has limite
'     d warranties.
'**************************************

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Function NTUser() As String
    Dim llngResultCode As Long
    Dim lstrBuffer As String * 255
    On Error Resume Next
    llngResultCode = WNetGetUser(vbNullString, lstrBuffer, 255) ' API Call
    NTUser = Trim(Replace(lstrBuffer, Chr(0), "") & "")
End Function