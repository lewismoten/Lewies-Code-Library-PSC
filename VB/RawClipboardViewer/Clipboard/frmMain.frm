VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Raw Clipboard Viewer"
   ClientHeight    =   5340
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox lstText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      ItemData        =   "frmMain.frx":0000
      Left            =   5280
      List            =   "frmMain.frx":0013
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox lstBytes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      ItemData        =   "frmMain.frx":0071
      Left            =   0
      List            =   "frmMain.frx":0093
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.OptionButton optView 
      Caption         =   "Data"
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optView 
      Caption         =   "Text"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuFormat_Bitmap 
         Caption         =   "Bitmap"
      End
      Begin VB.Menu mnuFormat_DIB 
         Caption         =   "DIB"
      End
      Begin VB.Menu mnuFormat_Metafile 
         Caption         =   "Metafile"
      End
      Begin VB.Menu mnuFormat_Files 
         Caption         =   "Files"
      End
      Begin VB.Menu mnuFormat_Link 
         Caption         =   "Link"
      End
      Begin VB.Menu mnuFormat_Palette 
         Caption         =   "Palette"
      End
      Begin VB.Menu mnuFormat_RTF 
         Caption         =   "RTF"
      End
      Begin VB.Menu mnuFormat_Text 
         Caption         =   "Text"
      End
      Begin VB.Menu mnuFormat_Other 
         Caption         =   "Other..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngMode As Long

Private Sub ReadClipBoard()
    
    Dim llngIndex As Long
    Dim lstrFormat As String
    
    lstBytes.Clear
    lstText.Clear
    frmOther.lstFormats.Clear
    
    Me.mnuFormat_Bitmap.Enabled = False
    Me.mnuFormat_DIB.Enabled = False
    Me.mnuFormat_Files.Enabled = False
    Me.mnuFormat_Link.Enabled = False
    Me.mnuFormat_Metafile.Enabled = False
    Me.mnuFormat_Other.Enabled = False
    Me.mnuFormat_Palette.Enabled = False
    Me.mnuFormat_RTF.Enabled = False
    Me.mnuFormat_Text.Enabled = False
    
    For llngIndex = -32768 To 32767
        If VB.Clipboard.GetFormat(llngIndex) Then
            lstrFormat = Right("     " & llngIndex, 6) & " "
            Select Case llngIndex
                Case vbCFBitmap
                    lstrFormat = lstrFormat & "Bitmap"
                    Me.mnuFormat_Bitmap.Enabled = True
                Case vbCFDIB
                    lstrFormat = lstrFormat & "DIB"
                    Me.mnuFormat_DIB.Enabled = True
                Case vbCFEMetafile
                    lstrFormat = lstrFormat & "Metafile"
                    Me.mnuFormat_Metafile.Enabled = True
                Case vbCFFiles
                    lstrFormat = lstrFormat & "Files"
                    Me.mnuFormat_Files.Enabled = True
                Case vbCFLink
                    lstrFormat = lstrFormat & "Link"
                    Me.mnuFormat_Link.Enabled = True
                Case vbCFMetafile
                    lstrFormat = lstrFormat & "Metafile"
                    Me.mnuFormat_Metafile.Enabled = True
                Case vbCFPalette
                    lstrFormat = lstrFormat & "Palette"
                    Me.mnuFormat_Palette.Enabled = True
                Case vbCFRTF
                    lstrFormat = lstrFormat & "RTF"
                    Me.mnuFormat_RTF.Enabled = True
                Case vbCFText
                    lstrFormat = lstrFormat & "Text"
                    Me.mnuFormat_Text.Enabled = True
                Case Else
                    Me.mnuFormat_Other.Enabled = True
            End Select
            frmOther.lstFormats.AddItem lstrFormat
            frmOther.lstFormats.ItemData(frmOther.lstFormats.NewIndex) = llngIndex
            DoEvents
        End If
    Next
End Sub


Private Sub cmdRefresh_Click()
    ReadClipBoard
End Sub

Private Sub Form_Load()
    ReadClipBoard
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
    Exit Sub
    lstBytes.Top = 0
    lstText.Top = 0
    
    lstBytes.Left = 0
    lstBytes.Width = Me.Width - Me.ScaleX(200, vbPixels, Me.ScaleMode)
    
    lstText.Left = lstBytes.Left + lstBytes.Width
    lstText.Width = Me.Width - lstBytes.Width
    
    lstBytes.Height = Me.Height
    lstText.Height = Me.Height
End Sub

Private Sub lstBytes_Click()
    If Not lstText.ListIndex = lstBytes.ListIndex Then _
        lstText.ListIndex = lstBytes.ListIndex
End Sub

Private Sub lstText_Click()
    If Not lstBytes.ListIndex = lstText.ListIndex Then _
        lstBytes.ListIndex = lstText.ListIndex
End Sub

Private Sub mnuFile_Exit_Click()
    End
End Sub

Private Sub mnuFile_Refresh_Click()
    ReadClipBoard
End Sub

Private Sub mnuFormat_Bitmap_Click()
    SetView ClipBoardConstants.vbCFBitmap
End Sub

Private Sub mnuFormat_DIB_Click()
    SetView ClipBoardConstants.vbCFDIB
End Sub

Private Sub mnuFormat_Files_Click()
    SetView ClipBoardConstants.vbCFFiles
End Sub

Private Sub mnuFormat_Link_Click()
    SetView ClipBoardConstants.vbCFLink
End Sub

Private Sub mnuFormat_Metafile_Click()
    SetView ClipBoardConstants.vbCFMetafile
End Sub

Private Sub mnuFormat_Other_Click()
    frmOther.Show vbModal, Me
End Sub

Private Sub mnuFormat_Palette_Click()
    SetView ClipBoardConstants.vbCFPalette
End Sub

Private Sub mnuFormat_RTF_Click()
    SetView ClipBoardConstants.vbCFRTF
End Sub

Private Sub mnuFormat_Text_Click()
    SetView ClipBoardConstants.vbCFText
End Sub
Public Sub SetView(ByRef plngClipboardFormat As Long)
    
    Dim lstrText As String
    Dim lobjPic As IPictureDisp
    
    mlngMode = plngClipboardFormat
    
    lstBytes.Clear
    lstText.Clear
    
    On Error Resume Next
    If optView(0).Value Then
        lstrText = VB.Clipboard.GetText(plngClipboardFormat)
    Else
        Set lobjPic = VB.Clipboard.GetData(plngClipboardFormat)
    End If
    
    lstText.ForeColor = vbWindowText
    lstBytes.ForeColor = vbWindowText
    If Err Then
        
        lstText.ForeColor = vbRed
        lstBytes.ForeColor = vbRed
        
        lstrText = Err.Number & " - " & Err.Description
        Err.Clear
        ShowText lstrText
        Exit Sub
        
    End If
    On Error GoTo 0
    
    If optView(0).Value Then
        ShowText lstrText
    Else
        ShowData lobjPic
    End If
End Sub

Private Sub ShowData(ByRef pobjPic As IPictureDisp)
    lstBytes.Visible = False
    On Error Resume Next
    lstText.AddItem "Handle: " & pobjPic.Handle
    lstText.AddItem "Height: " & pobjPic.Height
    lstText.AddItem "hPal: " & pobjPic.hPal
    lstText.AddItem "Type: " & pobjPic.Type
    lstText.AddItem "Width: " & pobjPic.Width
End Sub

Private Sub ShowText(ByRef pstrText)
    Dim llngIndex As Long
    Dim llngMaxIndex As Long
    Dim lstrChunk As String
    
    lstBytes.Visible = True
    lstText.Visible = True
    
    llngMaxIndex = Len(pstrText)
    For llngIndex = 1 To llngMaxIndex Step 16
        lstrChunk = Mid(pstrText, llngIndex, 16)
        lstText.AddItem lstrChunk
        lstBytes.AddItem ToBytes(lstrChunk)
    Next
End Sub
Private Function ToBytes(ByRef pstrString)
    Dim llngIndex As Long
    Dim llngMaxIndex As Long
    Dim lstrBytes
    llngMaxIndex = Len(pstrString)
    For llngIndex = 1 To llngMaxIndex
        lstrBytes = lstrBytes & Right("0" & Hex(Asc(Mid(pstrString, llngIndex, 1))), 2) & " "
    Next
    ToBytes = lstrBytes
End Function

Private Sub optView_Click(Index As Integer)
    SetView mlngMode
End Sub
