VERSION 5.00
Begin VB.Form frmKeySnooper 
   BackColor       =   &H8000000D&
   Caption         =   "KeySnooper Test Container"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmKeySnooper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToolbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   2745
      Width           =   4680
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Start Snooping"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.TextBox txtKeys 
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmKeySnooper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAction_Click()
    cmdAction.Enabled = False
    KeySnooper
End Sub

Private Sub cmdClear_Click()
    Cache = ""
    txtKeys.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
    txtKeys.Left = 0
    txtKeys.Top = 0
    txtKeys.Height = Me.ScaleHeight - picToolbar.ScaleHeight
    txtKeys.Width = Me.ScaleWidth
End Sub

Private Sub lblURL_Click()
    Call ShellExecute(Me.hwnd, "Open", "http://www.lewismoten.com", "", "", True)
End Sub
