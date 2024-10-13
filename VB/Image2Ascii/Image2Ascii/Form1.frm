VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Image to Ascii Converter"
   ClientHeight    =   1965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   0
      Width           =   1185
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuColumns 
         Caption         =   "Columns"
      End
      Begin VB.Menu mnuRows 
         Caption         =   "Rows"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
      End
   End
   Begin VB.Menu mnuWho 
      Caption         =   "Who?"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cc(30)
Dim rows
Dim columns
Dim t
Dim file

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Init
    Picture1 = LoadPicture(file)
    columns = 80
    rows = 25
    mnuColumns.Caption = "Columns " & columns
    mnuRows.Caption = "Rows " & rows
    
End Sub

Function GetLetter(Percent)

    Randomize
    
    Percent = (Percent * 100) \ 3

    If Percent > 30 Then Percent = 30

    i = Int(Rnd * Len(cc(Percent))) + 1
    
    c = Mid(cc(Percent), i, 1)
    
    GetLetter = c
    
End Function

Sub Init()

    file = App.Path & "\lewismoten.com.gif"
    cc(0) = " "
    cc(1) = " `"
    cc(2) = "`"
    cc(3) = "`'."
    cc(4) = "'."
    cc(5) = "'.-~!^,"
    cc(6) = "-~!^,"
    cc(7) = "&_"
    cc(10) = "()|\<>/"
    cc(11) = "*?"
    cc(8) = ":"
    cc(9) = "{};"
    cc(10) = "()|\<>/"
    cc(11) = "*?"
    cc(12) = "="""
    cc(13) = "il%+"
    cc(14) = "t17[]"
    cc(15) = "Ccjrv3"
    cc(16) = "IJoxz2"
    cc(17) = "Lu$"
    cc(18) = "OTYZ6"
    cc(19) = "GVenswy45"
    cc(20) = "ADFPbf0"
    cc(21) = "SUahp9"
    cc(22) = "Xdkm8"
    cc(23) = "QRq"
    cc(24) = "BEKWg"
    cc(25) = "@"
    cc(26) = "@HN"
    cc(27) = "HN"
    cc(28) = "#"
    cc(29) = "M#"
    cc(30) = "M"
End Sub

Private Sub Form_Resize()
'    Picture1.Width = Me.ScaleWidth
'    Picture1.Height = Me.ScaleHeight
End Sub

Private Sub mnuColumns_Click()
    columns = InputBox("How many columns?", "Settings", columns)
    mnuColumns.Caption = "Columns " & columns
End Sub

Private Sub mnuConvert_Click()
    t = ""
    For y = 0 To Picture1.Height
        For x = 0 To Picture1.Width
            c = Picture1.Point(x, y)
            h = Right("000000" & Hex(c), 6)
            r = CInt("&h" & Mid(h, 1, 2))
            g = CInt("&h" & Mid(h, 3, 2))
            b = CInt("&h" & Mid(h, 5, 2))
            a = (r + g + b) / 3
            n = (255 - a) / 255
            t = t & GetLetter(n)
        Next
        t = t & vbCrLf
        DoEvents
    Next
    Form2.Text1.Text = t
    Form2.Show vbModal
End Sub

Private Sub mnuLoad_Click()
    file = InputBox("Path to image file?", "Load", file)
    Picture1 = LoadPicture(file)
End Sub

Private Sub mnuRows_Click()
    rows = InputBox("How many rows?", "Settings", rows)
    mnuRows.Caption = "Rows " & rows
End Sub

Private Sub mnuWho_Click()
    Call ShellExecute(Me.hwnd, "Open", "http://www.lewismoten.com", "", "", True)
End Sub
