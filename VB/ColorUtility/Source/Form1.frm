VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColorUtil 
   Caption         =   "Color Utility"
   ClientHeight    =   2505
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   3120
      TabIndex        =   17
      Text            =   "Black"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton optSet 
      Caption         =   "RGB"
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.OptionButton optSet 
      Caption         =   "Name"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton optSet 
      Caption         =   "Hex"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSet 
      Caption         =   "Long"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdSetColor 
      Caption         =   "&Set Color"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtHex 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Text            =   "#000000"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtLong 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Text            =   "0"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Text            =   "0"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Text            =   "0"
      Top             =   1680
      Width           =   375
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   255
      Index           =   2
      LargeChange     =   16
      Left            =   600
      Max             =   255
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   255
      Index           =   1
      LargeChange     =   16
      Left            =   600
      Max             =   255
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   255
      Index           =   0
      LargeChange     =   16
      Left            =   600
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblSafeColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Safe Color"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "http://www.lewis.moten.com"
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   2100
      Width           =   3435
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape shpRGB 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   3000
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape shpRGB 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   3000
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape shpRGB 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   3000
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape shpColor 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3360
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   300
   End
   Begin VB.Menu mnuLoadImage 
      Caption         =   "&Load Image"
   End
   Begin VB.Menu mnuSystemPalette 
      Caption         =   "&System Palette"
   End
End
Attribute VB_Name = "frmColorUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gRGB(2) As Byte
Dim WebColors() As Variant
Dim WebColorNames() As Variant
Function HexNum(HexChar As String) As Byte
    If IsNumeric(HexChar) Then
        HexNum = CInt(HexChar)
    Else
        Select Case LCase(HexChar)
            Case "a"
                HexNum = 10
            Case "b"
                HexNum = 11
            Case "c"
                HexNum = 12
            Case "d"
                HexNum = 13
            Case "e"
                HexNum = 14
            Case "f"
                HexNum = 15
            Case Else
                MsgBox "Your Hex Number had an invalid character """ & HexChar & """, it will be changed to zero."
                HexNum = 0
        End Select
    End If
End Function

Private Sub cboName_Click()
    'cmdSetColor_Click
End Sub

Private Sub cboName_GotFocus()
    optSet(2).Value = True
End Sub

Private Sub cmdSetColor_Click()
    Dim Count As Integer
    For Count = 0 To 3
        If optSet(Count).Value = True Then Exit For
    Next
    Select Case Count
        Case 0 'Long
            If IsNumeric(txtLong) And txtLong <> "" Then
                If txtLong <= 16777215 And txtLong >= 0 Then
                    LongToRGB txtLong
                Else
                    MsgBox "Please enter a number from 0 To 16777215 in the ""Long"" field."
                End If
            Else
                MsgBox "You must enter a Number for the ""Long"" field."
            End If
        Case 1 'Hex
            Select Case InStr(txtHex, "#")
                Case 0
                    txtHex = "#" & txtHex
                Case 1
                    'Everything is good.
                Case Else
                    MsgBox "It appears that your # sign is not at the beginning of your Hex Number"
                    Exit Sub
            End Select
            If Len(txtHex) <> 7 Then
                MsgBox "You must have 6 characters following a pound sign ""#"" in your Hex string."
            End If
            For Count = 0 To 2
                gRGB(Count) = _
                (HexNum(Mid(txtHex, (Count * 2) + 2, 1)) * 16) _
                + HexNum(Mid(txtHex, (Count * 2) + 3, 1))
            Next
        Case 2 'Name
            LongToRGB WebColors(cboName.ListIndex)
        Case 3 'RGB
            For Count = 0 To 2
                gRGB(Count) = CByte(txtRGB(Count))
            Next
    End Select
    SetAllFromRGB
End Sub

Private Sub Form_Load()
    Dim Count As Long
    WebColors = Array(0, 16777215, 8421504, 12632256, 255, 32768, 16711680, 65535, 8388736, 32896, 8388608, _
        16776960, 65280, 128, 8421376, 16711935, 25600, 7346457, 6908265, 9470064, 10061943, _
        9109504, 13434880, 9145088, 16760576, 13749760, 10156544, 8388352, 16776960, 16748574, _
        11186720, 2263842, 5737262, 5197615, 3329330, 7451452, 13688896, 14772545, 11829830, _
        9125192, 13422920, 8519755, 3107669, 10526303, 15570276, 11193702, 13458026, 2330219, _
        15624315, 64636, 65407, 13959039, 15453831, 16436871, 14822282, 139, 9109643, 1262987, _
        9157775, 9498256, 14381203, 13828244, 10025880, 13382297, 3329434, 2970272, 2763429, _
        11119017, 15128749, 3145645, 15658671, 14599344, 15130800, 2237106, 755384, 13850042, _
        9408444, 7059389, 8721863, 6053069, 4163021, 1993170, 9221330, 13882323, 14204888, _
        14053594, 2139610, 9662683, 3937500, 14474460, 14524637, 8894686, 16777184, 16443110, _
        8034025, 15631086, 15136253, 16711935, 9639167, 17919, 4678655, 11823615, 5275647, _
        36095, 8036607, 42495, 12695295, 13353215, 55295, 12180223, 11394815, 11920639, 12903679, _
        14804223, 13495295, 14020607, 16118015, 15660543, 14481663, 13499135, 15792895, _
        16448255, 14745599, 15794175 _
        )
    WebColorNames = Array("Black", "White", "Gray", "Silver", "Red", "Green", "Blue", "Yellow", "Purple", _
        "Olive", "Navy", "Aqua", "Lime", "Maroon", "Teal", "Fuchsia", "DarkGreen", _
        "MidnightBlue", "DimGray", "SlateGray", "LightSlateGray", "DarkBlue", "MediumBlue", _
        "DarkCyan", "DeepSkyBlue", "DarkTurquoise", "MediumSpringGreen", "SpringGreen", _
        "Aqua", "DodgerBlue", "LightSeaGreen", "ForestGreen", "SeaGreen", "DarkSlateGray", _
        "LimeGreen", "MediumSeaGreen", "Turquoise", "RoyalBlue", "SteelBlue", "DarkSlateBlue", _
        "MediumTurquoise", "Indigo", "DarkOliveGreen", "DarkOliveGreen", "CornFlowerBlue", _
        "MediumAquamarine", "SlateBlue", "OliveDrab", "MediumSlateBlue", "LawnGreen", _
        "Chartreuse", "Aquamarine", "SkyBlue", "LightSkyBlue", "BlueViolet", "DarkRed", _
        "DarkMagenta", "SaddleBrown", "DarkSeaGreen", "LightGreen", "MediumPurple", "DarkViolet", _
        "PaleGreen", "DarkOrchid", "YellowGreen", "Sienna", "Brown", "DarkGray", "LightBlue", _
        "GreenYellow", "PaleTurquoise", "LightSteelBlue", "PowderBlue", "FireBrick", "DarkGoldenrod", _
        "MediumOrchid", "RosyBrown", "Darkkhaki", "MediumVioletRed", "IndianRed", "Peru", _
        "Chocolate", "Tan", "LightGrey", "Thistle", "Orchid", "Goldenrod", "PaleVioletRed", "Crimson", _
        "Gainsboro", "Plum", "Burlywood", "LightCyan", "Lavender", "DarkSalmon", "Violet", _
        "Oldlace", "Fuchsia", "DeepPink", "OrangeRed", "Tomato", "HotPink", "Coral", "DarkOrange", _
        "LightSalmon", "Orange", "LightPink", "Pink", "Gold", "Peachpuff", "NavajoWhite", "Moccasin", _
        "Bisque", "MistyRose", "BlanchedAlmond", "PapayaWhip", "LavenderBlush", "Seashell", _
        "Cornsilk", "LemonChiffon", "FloralWhite", "Snow", "LightYellow", "Ivory" _
        )
    
    For Count = 0 To UBound(WebColorNames)
        cboName.AddItem WebColorNames(Count)
    Next

End Sub

Private Sub hsbRGB_Change(Index As Integer)
    gRGB(Index) = hsbRGB(Index).Value
    SetAllFromRGB
End Sub
Public Sub SetAllFromRGB()
    Dim Count As Byte
    Dim OKSoFar As Boolean
    OKSoFar = True
    txtHex = "#"
    For Count = 0 To 2
        txtRGB(Count) = gRGB(Count)
        If Len(Hex(gRGB(Count))) = 1 Then
            txtHex = txtHex & "0"
        End If
        txtHex = txtHex & Format(Hex(gRGB(Count)), "<") '"0#"
    Next
    For Count = 0 To 2
        hsbRGB(Count) = gRGB(Count)
    Next
    shpRGB(0).FillColor = RGB(gRGB(0), 0, 0)
    shpRGB(1).FillColor = RGB(0, gRGB(1), 0)
    shpRGB(2).FillColor = RGB(0, 0, gRGB(2))
    shpColor.FillColor = RGB(gRGB(0), gRGB(1), gRGB(2))
    txtLong = RGB(gRGB(0), gRGB(1), gRGB(2))
    cboName.Text = ""
    cboName.ListIndex = -1
    For Count = 0 To UBound(WebColors)
        If WebColors(Count) = txtLong Then
            cboName.ListIndex = Count
            Exit For
        End If
    Next
    'Test to see if it is web-color pallette safe.
    For Count = 0 To 2
        If gRGB(Count) <> 0 Then
            If Len(Hex(gRGB(Count))) = 1 Then
                OKSoFar = False
                Exit For
            Else
                If Mid(Hex(gRGB(Count)), 1, 1) <> Mid(Hex(gRGB(Count)), 2, 1) Then
                    OKSoFar = False
                    Exit For
                End If
            End If
        End If
    Next
    If OKSoFar Then
        Me.lblSafeColor.Caption = "Safe Color"
    Else
        Me.lblSafeColor.Caption = ""
    End If
                
End Sub

Sub LongToRGB(ByVal LongValue As Long)
    'Dim lvRGB(2) As Byte
    Dim Count As Integer
    Dim Temp As Long
    
    ldLongIP = CDbl(txtLongIP)
    
    For Count = 2 To 0 Step -1
        Temp = Int(LongValue / (256 ^ Count))
        gRGB(Count) = Temp
        LongValue = LongValue - (Temp * (256 ^ Count))
        'MsgBox (2 - Count) & " = " & lvRGB(2 - Count)
    Next
    
End Sub
Private Sub hsbRGB_Scroll(Index As Integer)
    hsbRGB_Change Index
End Sub

Private Sub mnuLoadImage_Click()
    On Error Resume Next
    CommonDialog1.Filter = "Picture |*.bmp;*.gif;*.ico;*.jpg|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    If Err Then Exit Sub
    frmColorPick.Picture1.Picture = LoadPicture(CommonDialog1.FileName)
    frmColorPick.Show
End Sub

Private Sub mnuSystemPalette_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    If Err Then Exit Sub
    LongToRGB CommonDialog1.Color
    SetAllFromRGB
End Sub

Private Sub txtHex_GotFocus()
    optSet(1).Value = True
End Sub

Private Sub txtLong_GotFocus()
    optSet(0).Value = True
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)
    optSet(3).Value = True
End Sub
