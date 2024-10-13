VERSION 5.00
Begin VB.Form frmTextConverter 
   Caption         =   "Text Converter"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5460
      TabIndex        =   1
      Top             =   1800
      Width           =   5460
      Begin VB.CommandButton cmdWho 
         Caption         =   "Who?"
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtVariable 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "lstrResult"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdVBS 
         Caption         =   "&VBS"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdSQL 
         Caption         =   "SQL"
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdJS 
         Caption         =   "JS"
         Height          =   495
         Left            =   3600
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Variable Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtText 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmTextConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdVBS_Click()
    Dim lstrLinesAry As Variant
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrLine As String
    lstrLinesAry = Split(txtText.Text, vbCrLf)
    llngMaxIndex = UBound(lstrLinesAry)
    For llngIndex = 0 To llngMaxIndex
        lstrLine = lstrLinesAry(llngIndex)
        lstrLine = Replace(lstrLine, """", """""")
        lstrLine = Replace(lstrLine, vbTab, """ & vbTab & """)
        lstrLine = txtVariable & " = " & txtVariable & " & """ & lstrLine & """"""
        lstrLinesAry(llngIndex) = lstrLine
    Next
    txtText.Text = "Dim " & txtVariable.Text & " As String" & vbCrLf & _
        Join(lstrLinesAry, vbCrLf)
End Sub

Private Sub cmdJS_Click()
    Dim lstrLinesAry As Variant
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrLine As String
    lstrLinesAry = Split(txtText.Text, vbCrLf)
    llngMaxIndex = UBound(lstrLinesAry)
    For llngIndex = 0 To llngMaxIndex
        lstrLine = lstrLinesAry(llngIndex)
        lstrLine = Replace(lstrLine, "\", "\\")
        lstrLine = Replace(lstrLine, "'", "\'")
        lstrLine = Replace(lstrLine, vbTab, "\t")
        lstrLine = txtVariable & " += '" & lstrLine & "'"
        lstrLinesAry(llngIndex) = lstrLine
    Next
    txtText.Text = "var " & txtVariable.Text & " = new String('')" & vbCrLf & _
        Join(lstrLinesAry, vbCrLf)
End Sub

Private Sub cmdSQL_Click()
    Dim lstrLinesAry As Variant
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrLine As String
    lstrLinesAry = Split(txtText.Text, vbCrLf)
    llngMaxIndex = UBound(lstrLinesAry)
    For llngIndex = 0 To llngMaxIndex
        lstrLine = lstrLinesAry(llngIndex)
        lstrLine = Replace(lstrLine, "'", "''")
        lstrLine = "SET " & txtVariable.Text & " = " & txtVariable.Text & " + '" & lstrLine & "'"
        lstrLinesAry(llngIndex) = lstrLine
    Next
    txtText.Text = "DECLARE " & txtVariable.Text & " VARCHAR(8000)" & vbCrLf & _
        "SET " & txtVariable.Text & " = ''" & vbCrLf & _
        Join(lstrLinesAry, vbCrLf)
End Sub

Private Sub cmdWho_Click()

    Call ShellExecute(Me.hwnd, "Open", "http://www.lewismoten.com", "", "", True)
    
End Sub

Private Sub Form_Resize()
    
    If Me.Width < 5580 Then Me.Width = 5580
    If Me.Height < 2940 Then Me.Height = 2940
    
    txtText.Width = Me.ScaleWidth
    txtText.Height = Me.ScaleHeight - picBar.Height
    txtText.Left = 0
    txtText.Top = 0

End Sub
