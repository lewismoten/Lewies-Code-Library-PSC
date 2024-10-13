VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Script.Decoder"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   3480
      TabIndex        =   1
      Top             =   1755
      Width           =   3480
      Begin VB.CommandButton Command1 
         Caption         =   "Decode"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "www.lewismoten.com"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim s As New Class1
Me.Text1.Text = s.DecodeScriptFile("Script", Me.Text1.Text, 0, "")
Set s = Nothing
End Sub


Private Sub Form_Resize()
With Text1
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight - Picture1.Height
.Left = 0
.Top = 0
End With
End Sub
