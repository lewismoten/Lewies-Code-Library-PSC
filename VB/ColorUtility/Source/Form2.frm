VERSION 5.00
Begin VB.Form frmColorPick 
   Caption         =   "Color Picker"
   ClientHeight    =   2085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    With Picture1
        .Left = 0
        .Top = 0
        .Width = Width
        .Height = Height
    End With
    With picColor
        .Left = Width - .Width - .Width
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmColorUtil.LongToRGB Point(X, Y)
End Sub

Private Sub Form_Resize()
    Form_Load
End Sub

Private Sub Picture1_Click()
    frmColorUtil.LongToRGB picColor.BackColor
    frmColorUtil.SetAllFromRGB
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picColor.BackColor = Picture1.Point(X, Y)
End Sub
