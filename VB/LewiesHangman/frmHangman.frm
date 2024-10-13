VERSION 5.00
Begin VB.Form frmHangman 
   Caption         =   "Lewie's Hangman"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuess 
      Caption         =   "Guess"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remaining"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   1080
         TabIndex        =   26
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   840
         TabIndex        =   25
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   24
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Created by Lewis Moten           http://www.lewismoten.com/"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Line manLeg2 
      X1              =   2400
      X2              =   2520
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line manLeg1 
      X1              =   2160
      X2              =   2040
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line manArm2 
      X1              =   2400
      X2              =   2640
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line manArm1 
      X1              =   1920
      X2              =   2160
      Y1              =   720
      Y2              =   480
   End
   Begin VB.Shape manBody 
      Height          =   495
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape manHead 
      Height          =   255
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line9 
      X1              =   3120
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line8 
      X1              =   3240
      X2              =   3360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line7 
      X1              =   3360
      X2              =   3360
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Line Line6 
      X1              =   3240
      X2              =   3240
      Y1              =   1680
      Y2              =   1800
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   3120
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line4 
      X1              =   2760
      X2              =   3120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2280
      Y1              =   120
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblAnswer 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   4455
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Phrase As String
Dim Wrong As Integer
Dim PhraseAry As Variant
Private Sub LoadPhrases()
    Dim llngFileNum As Long
    Dim llngFileLen As Long
    
    lstrFileName = App.Path & "\Phrases.txt"
    
    On Error Resume Next
    llngFileLen = 0
    llngFileLen = FileLen(lstrFileName)
    
    If Err Then
        Err.Clear
        MsgBox "The file """ & lstrFileName & """ could not be found."
        End
    End If
    
    If llngFileLen = 0 Then Exit Sub
    On Error GoTo 0
    llngFileNum = FreeFile
    Open lstrFileName For Input As #llngFileNum
    Phrases = Input(llngFileLen, #llngFileNum)
    Close #lngFileNum
    
    PhraseAry = Split(Phrases, vbCrLf)
    
End Sub

Private Sub cmdGuess_Click()
    If InputBox("What is the answer?", "Guess", lblAnswer.Caption) = Phrase Then
        MsgBox "You Won!", vbExclamation, "Yay!"
        Call ResetLetters
        Call NewPhrase
        Call ShowPhrase
    Else
        MsgBox "I don't think so!", vbCritical, "Wrong Answer"
    End If
End Sub

Private Sub Form_Load()
    Call LoadPhrases
    Call ResetLetters
    Call NewPhrase
    Call ShowPhrase
End Sub

Private Sub lblLetter_Click(Index As Integer)
    If lblLetter(Index).Caption = "" Then Exit Sub
    If InStr(1, Phrase, lblLetter(Index).Caption, vbTextCompare) = 0 Then
        lblLetter(Index).Caption = ""
        Call ShowPhrase
        Call ShowPhrase
        Wrong = Wrong + 1
        Select Case Wrong
            Case 1: manHead.Visible = True
            Case 2: manBody.Visible = True
            Case 3: manArm1.Visible = True
            Case 4: manArm2.Visible = True
            Case 5: manLeg1.Visible = True
            Case 6
                manLeg2.Visible = True
                MsgBox "You Lost!", vbCritical, "Oops!"
                Call ResetLetters
                Call NewPhrase
                Call ShowPhrase
        End Select
        Exit Sub
    End If
    
    lblLetter(Index).Caption = ""
    Call ShowPhrase
    
    If lblAnswer.Caption = Phrase Then
        MsgBox "You Won!", vbExclamation, "Yay!"
        Call ResetLetters
        Call NewPhrase
        Call ShowPhrase
    End If
    
End Sub

Private Sub ResetLetters()
    
    Dim Index As Integer
    For Index = 0 To lblLetter.Count - 1
        lblLetter(Index).Caption = Chr(Asc("A") + Index)
    Next
    manArm1.Visible = False
    manArm2.Visible = False
    manBody.Visible = False
    manHead.Visible = False
    manLeg1.Visible = False
    manLeg2.Visible = False
    
End Sub

Private Sub NewPhrase()
    Dim Index
    Randomize
    Index = Int(Rnd * (UBound(PhraseAry) + 1))
    Phrase = PhraseAry(Index)
    Wrong = 0
End Sub

Private Sub ShowPhrase()
    Dim Display As String
    Dim Index
    Display = Phrase
    For Index = 0 To lblLetter.Count - 1
        Display = Replace(Display, lblLetter(Index).Caption, "_", 1, -1, vbTextCompare)
    Next
    lblAnswer.Caption = Display
End Sub
