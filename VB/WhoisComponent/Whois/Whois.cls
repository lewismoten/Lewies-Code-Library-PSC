VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Whois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT
' WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
' OF MERCHANTABILITY AND/OR FITNESS FOR A  PARTICULAR
' PURPOSE.

' Copyright 2000. All rights reserved.
' Lewis Moten
' http://www.lewismoten.com/
' email: lewis@moten.com

Option Explicit

Private WithEvents mObjWinsock As MSWinsockLib.Winsock
Attribute mObjWinsock.VB_VarHelpID = -1

Private mStrWhoisServerAddress As String
Private mLngWhoisServerPort As Long
Private mBolRemoveCopyright As Boolean
Private mStrEndCopyright As String
Private mStrWhois As String
Private mBolDone As Boolean
Private mDtmStart As Date
Private mStrAvailableString As String
Private mLngTimeout As Long
Private mStrDomain As String

Public Property Get AvailableString() As Variant
    
    AvailableString = mStrAvailableString
    
End Property

Public Property Let AvailableString(ByRef pStrAvailable)
    
    mStrAvailableString = pStrAvailable
    
End Property

Public Property Get EndCopyright() As Variant
    
    EndCopyright = mStrEndCopyright

End Property

Public Property Let EndCopyright(ByRef pStrEndText As Variant)
    
    mStrEndCopyright = pStrEndText
    
End Property

Public Property Get IsAvailable() As Variant

    IsAvailable = Not InStr(1, mStrWhois, mStrAvailableString, vbTextCompare) = 0

End Property

Public Sub Query(ByRef pStrDomain As Variant)
    
    mStrWhois = ""
    mBolDone = False
    mDtmStart = Now()
    mStrDomain = pStrDomain
    
    mObjWinsock.Connect mStrWhoisServerAddress, mLngWhoisServerPort
    
    While Not mBolDone And DateDiff("s", mDtmStart, Now()) < mLngTimeout
        DoEvents
    Wend
    
    mObjWinsock.Close
    
End Sub

Public Property Get RemoveCopyright() As Variant
    
    RemoveCopyright = mBolRemoveCopyright

End Property

Public Property Let RemoveCopyright(ByRef pBolRemove As Variant)
    
    mBolRemoveCopyright = pBolRemove

End Property

Public Property Get Timeout() As Variant
    
    Timeout = mLngTimeout

End Property

Public Property Let Timeout(ByRef pLngSeconds As Variant)
    
    mLngTimeout = pLngSeconds

End Property

Public Property Get Whois() As Variant
    
    Dim lLngCopyStart
    
    If RemoveCopyright Then
        
        lLngCopyStart = InStr(1, mStrWhois, mStrEndCopyright, vbTextCompare)
        
        If lLngCopyStart = 0 Then
            Whois = mStrWhois
        Else
            Whois = Mid(mStrWhois, lLngCopyStart + Len(mStrEndCopyright))
        End If
    
    Else
        
        Whois = mStrWhois
    
    End If
    
End Property

Public Property Get WhoisHtml() As Variant
    
    WhoisHtml = "<PRE>" & Whois & "</PRE>"

End Property

Public Property Get WhoisServerAddress() As Variant
    
    WhoisServerAddress = mStrWhoisServerAddress

End Property

Public Property Let WhoisServerAddress(ByRef pStrAddress As Variant)
    
    mStrWhoisServerAddress = pStrAddress

End Property

Public Property Get WhoisServerPort() As Variant
    
    WhoisServerPort = mLngWhoisServerPort

End Property

Public Property Let WhoisServerPort(ByRef pLngPort As Variant)
    
    mLngWhoisServerPort = pLngPort

End Property

Private Sub Class_Initialize()
    
    Set mObjWinsock = New MSWinsockLib.Winsock
    
    mStrWhoisServerAddress = "whois.networksolutions.com"
    mLngWhoisServerPort = 43
    mBolRemoveCopyright = False
    mStrEndCopyright = "policy."
    mStrWhois = "[Error: You must query for the domain first]"
    mBolDone = True
    mLngTimeout = 10
    mStrAvailableString = "no match"
    
End Sub

Private Sub Class_Terminate()
    
    Set mObjWinsock = Nothing

End Sub

Private Sub mObjWinsock_Close()
    
    mBolDone = True
    mStrWhois = mStrWhois

End Sub

Private Sub mObjWinsock_Connect()
    
    mObjWinsock.SendData mStrDomain & vbCrLf

End Sub

Private Sub mObjWinsock_DataArrival(ByVal bytesTotal As Long)
    
    Dim lStrData As String
    
    mObjWinsock.GetData lStrData, vbString
    
    If Not mStrWhois = "" Then mStrWhois = mStrWhois & vbCrLf
    
    mStrWhois = mStrWhois & lStrData
    
End Sub










