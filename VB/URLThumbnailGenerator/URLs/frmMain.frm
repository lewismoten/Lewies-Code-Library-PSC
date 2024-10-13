VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "URL Thumbnail Generator"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSaveAs 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "C:\ScreenShot.jpg"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.ComboBox cboScreenShot 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Text            =   "120x100"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox cboBrowserWindow 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "640x480"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox picScreenShot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.lewismoten.com"
      Top             =   480
      Width           =   5295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Create Thumbnail"
      Height          =   405
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblLink 
      Caption         =   "http://www.lewismoten.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "Utility Created By Lewis Moten"
      Height          =   195
      Left            =   3240
      TabIndex        =   11
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label Label5 
      Caption         =   "Preview:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "URL:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Save As:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Screenshot Size:"
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Browser Window:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public BrowserWidth As Long
Public BrowserHeight As Long

Private Sub cboScreenShot_Click()
    Dim Resolution As String
    Resolution = cboScreenShot.Text
    If InStr(1, Resolution, "x") = 0 Then Exit Sub
    If Len(Resolution) = InStr(1, Resolution, "x") Then Exit Sub
    picScreenShot.Width = Mid(Resolution, 1, InStr(1, Resolution, "x") - 1)
    picScreenShot.Height = Mid(Resolution, InStr(1, Resolution, "x") + 1)
End Sub

Private Sub cboScreenShot_KeyUp(KeyCode As Integer, Shift As Integer)
    cboScreenShot_Click
End Sub

Private Sub cmdGo_Click()
    Dim Resolution As String
    If txtURL.Text = "" Then
        MsgBox "Invalid URL"
        Exit Sub
    End If
    
    Resolution = cboBrowserWindow.Text
    If InStr(1, Resolution, "x") = 0 Then
        MsgBox "Must define screenshot size for image (100x100)"
        Exit Sub
    End If
    
    Resolution = cboBrowserWindow.Text
    If InStr(1, Resolution, "x") = 0 Then
        MsgBox "Must define screenshot size for browser window (640x480)"
        Exit Sub
    End If
    
    BrowserWidth = Mid(Resolution, 1, InStr(1, Resolution, "x") - 1)
    BrowserHeight = Mid(Resolution, InStr(1, Resolution, "x") + 1)
    
    frmBrowser.Show
    'Resize the form
    frmBrowser.ScaleMode = vbPixels
    frmBrowser.Width = Me.ScaleX(frmMain.BrowserWidth, vbPixels, vbTwips)
    frmBrowser.Height = Me.ScaleY(frmMain.BrowserHeight, vbPixels, vbTwips)
    frmBrowser.Left = 1
    frmBrowser.Top = 1
    ' resize web browser
    With frmBrowser.WebBrowser1
        .Width = frmMain.BrowserWidth
        .Height = frmMain.BrowserHeight
        .Top = 1
        .Left = 1
    End With
    ' Navigate to URL
    frmBrowser.WebBrowser1.Navigate2 frmMain.txtURL.Text
   

End Sub

Private Sub Form_Load()
    cboBrowserWindow.AddItem "640x480"
    cboBrowserWindow.AddItem "800x600"
    cboBrowserWindow.AddItem "1024x768"
    cboBrowserWindow.AddItem "1280x1024"
    cboBrowserWindow.Text = "640x480"

    cboScreenShot.AddItem "120x100"
    cboScreenShot.AddItem "320x200"
    cboScreenShot.Text = "120x100"
    picScreenShot.Width = 200
    picScreenShot.Height = 140
    'cboScreenShot.Text = picScreenShot.Width & "x" & picScreenShot.Height
    
    txtSaveAs.Text = App.Path & "\ScreenShot.jpg"
End Sub

Private Sub lblLink_Click()
    Call ShellExecute(Me.hwnd, "Open", "http://www.lewismoten.com", "", "", True)
End Sub
