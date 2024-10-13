VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   1920
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnFirstHit As Boolean


Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    
Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Sub Form_Load()
    blnFirstHit = True
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    With WebBrowser1
            If .ReadyState = READYSTATE_INTERACTIVE Then Exit Sub
            If Not .ReadyState = READYSTATE_COMPLETE Then Exit Sub
            If .Busy = True Then Exit Sub
    End With

    
    ' Set ratio
'    Dim Ratio As Double
'    Ratio = frmMain.picScreenShot.Width / WebBrowser1.Width
'    frmMain.picScreenShot.Height = WebBrowser1.Height * Ratio
    
    ' Resize and paste to main form
    StretchBlt _
        frmMain.picScreenShot.hdc, _
        0, _
        0, _
        frmMain.picScreenShot.Width, _
        frmMain.picScreenShot.Height, _
        GetDC(WebBrowser1.Parent.hwnd), _
        0, _
        0, _
        WebBrowser1.Width, _
        WebBrowser1.Height, _
        SRCCOPY
    
    ' Save screenshot
    
    SavePicture frmMain.picScreenShot.Image, frmMain.txtSaveAs.Text
    
    'WebBrowser1.Navigate2 "about:blank"
    
    Unload Me

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    With WebBrowser1
        While Not ( _
            .ReadyState = READYSTATE_INTERACTIVE _
            Or .ReadyState = READYSTATE_COMPLETE _
            Or .Busy = True)
            DoEvents
        Wend
    End With
    Timer1.Enabled = True
End Sub

