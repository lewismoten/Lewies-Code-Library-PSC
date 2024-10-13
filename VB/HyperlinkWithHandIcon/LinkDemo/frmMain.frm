VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Link Demonstration"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "This demonstration comes with the file HAND.CUR for you to use with your other applications."
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Visit http://www.lewismoten.com"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2325
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API call to execute commands with the windows shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    ' Format the label to look like a URL
    DisplayAsURL Label1
End Sub

Private Sub Label1_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.lewismoten.com"
End Sub

Private Sub BrowseTo(ByRef pstrURL As String)
    ' Opens users default web browser and navigates to the selected URL
    Call ShellExecute(Me.hwnd, "Open", pstrURL, "", "", True)
End Sub

Private Sub DisplayAsURL(ByRef Link As VB.Label)
    ' Changes a link to look like a URL
    Link.Font.Underline = True
    Link.ForeColor = vbBlue
    Link.MousePointer = vbCustom
    Link.MouseIcon = LoadPicture(App.Path & "\Hand.cur")
End Sub
