VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Alpha Blend Demo"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1920
   End
   Begin VB.PictureBox picFinal 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.Label lblURL 
         BackStyle       =   0  'Transparent
         Caption         =   "www.lewismoten.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   5
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "Created by Lewis Moten"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   0
         Width           =   1710
      End
   End
   Begin VB.PictureBox picAlpha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picOverlay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "<-- Base"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "<-- Alpha"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "<-- Overlay"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This program was designed to help me create multiple sets of tiles for an RPG game.
' When designing land, I have found that I was basically creating the same kind of
' blends for all other seemless land combinations.  This program is the first step to
' automating the process.  It is not ment for an end user due to its long execution time.
' It is simply a utility

' Although the method is slow, it is a working method.  I could not find API's to do
' precisely what I wanted to do.  This program combines one image with another, using a
' third image to determine how much it should blend for each pixel (and each primary color).

' It is best to work with alpha maps that are in gray scale when beginning to use this program.
' From there, you may wish to experment with techniques involving 3 alpha channels - 1 for each
' color of Red, Green, and Blue

' If anyone has found a proven method to do this with API calls and without using code from another
' language other then VB, then I would be interested to hear your solution to the problem regarding
' speed.

' Base: The starting image that the overlay will be applied to
' Overlay: Image being laid over the base image
' Alpha: the opacity map on a pixel-per-pixel bases with three color channels
    
    ' Black = 0%
    ' White = 100%
    
    ' 0%   = Base 100%, Overlay 0%
    ' 10%  = Base 90%,  Overlay 10%
    ' 20%  = Base 80%,  Overlay 20%
    ' 30%  = Base 70%,  Overlay 30%
    ' 40%  = Base 60%,  Overlay 40%
    ' 50%  = Base 50%,  Overlay 50%
    ' 60%  = Base 40%,  Overlay 60%
    ' 70%  = Base 30%,  Overlay 70%
    ' 80%  = Base 20%,  Overlay 80%
    ' 90%  = Base 10%,  Overlay 90%
    ' 100% = Base 0%,   Overlay 100%

Dim DoAnother As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Load()
    Timer1.Enabled = True
    picFinal.Top = 0
    picFinal.Left = 0
End Sub

Function Blend(ByRef base As Long, ByRef alpha As Long, ByRef overlay As Long) As Long
    
    Dim BaseH As String
    Dim AlphaH As String
    Dim OverlayH As String
    
    Dim BaseR As Long
    Dim BaseG As Long
    Dim BaseB As Long
    
    Dim AlphaR As Long
    Dim AlphaG As Long
    Dim AlphaB As Long
    
    Dim OverlayR As Long
    Dim OverlayG As Long
    Dim OverlayB As Long
    
    Dim NewR As Long
    Dim NewG As Long
    Dim NewB As Long
    
    ' convert to hex values
    BaseH = Right("00000" & Hex(base), 6)
    AlphaH = Right("00000" & Hex(alpha), 6)
    OverlayH = Right("00000" & Hex(overlay), 6)

    ' Notice: Values are stored as BBGGRR (blue, green, red) when
    ' converted to a hex string - This is opposite of HTML colors
    ' stored as RRGGBB

    ' Pull out RGB values for base
    BaseB = CInt("&h" & Left(BaseH, 2))
    BaseG = CInt("&h" & Mid(BaseH, 3, 2))
    BaseR = CInt("&h" & Right(BaseH, 2))
    
    ' Pull out RGB values for alpha mask
    AlphaB = CInt("&h" & Left(AlphaH, 2))
    AlphaG = CInt("&h" & Mid(AlphaH, 3, 2))
    AlphaR = CInt("&h" & Right(AlphaH, 2))
    
    ' Pull out RGB values for overlay
    OverlayB = CInt("&h" & Left(OverlayH, 2))
    OverlayG = CInt("&h" & Mid(OverlayH, 3, 2))
    OverlayR = CInt("&h" & Right(OverlayH, 2))

    ' adjust base values based on alpha map opacity applied to
    ' difference between overlay and base maps
    NewR = BaseR + ((AlphaR / 255) * (OverlayR - BaseR))
    NewG = BaseG + ((AlphaG / 255) * (OverlayG - BaseG))
    NewB = BaseB + ((AlphaB / 255) * (OverlayB - BaseB))
    
    ' Make sure new values are within byte range
    If NewR > 255 Then NewR = 255
    If NewG > 255 Then NewG = 255
    If NewB > 255 Then NewB = 255
    
    If NewR < 0 Then NewR = 0
    If NewG < 0 Then NewG = 0
    If NewB < 0 Then NewB = 0
    
    ' convert values to rgb and return results
    Blend = RGB(NewR, NewG, NewB)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub lblURL_Click()
    Call ShellExecute(Me.hwnd, "Open", "http://www.lewismoten.com", "", "", True)
End Sub

Private Sub picFinal_Click()
    DoAnother = True
End Sub

Private Sub Timer1_Timer()
    Dim baseRGB As Long
    Dim alphaRGB As Long
    Dim overlayRGB As Long
    Dim newRGB As Long
    Dim x As Long
    Dim y As Long
    Dim w As Long
    Dim h As Long
    
    Dim DemoBase As Variant
    Dim Pic1 As String
    Dim Pic2 As String
    
    ' Disable Timer
    Timer1.Enabled = False
    
    ' Choose a random base tile to demonstrate
    DemoBase = Array("Volcano", "CoastalWaters", "DeepGrass", "DirtPath", "Grass", "Lava", "Mountains", "Ocean", "Sand")
    Randomize
    
    Pic1 = DemoBase(Int(Rnd * UBound(DemoBase)))
    Pic2 = DemoBase(Int(Rnd * UBound(DemoBase)))
    
    ' Make sure pictures are not the same
    While Pic1 = Pic2
        Pic2 = DemoBase(Int(Rnd * UBound(DemoBase)))
    Wend
    
    Me.Caption = "Alpha Blend Demo - " & Pic1 & ":" & Pic2
    ' Load Pictures
    picOverlay.Picture = LoadPicture(App.Path & "\BaseTiles-" & Pic1 & ".gif")
    picAlpha.Picture = LoadPicture(App.Path & "\AlphaMap-Ocean.gif")
    picBase.Picture = LoadPicture(App.Path & "\BaseTiles-" & Pic2 & ".gif")
    
    ' Copy & Tile picBase onto picFinal to be same height/width
    ' as picOverlay
    
    picFinal.Width = picAlpha.Width
    picFinal.Height = picAlpha.Height
    Me.Width = Me.ScaleX(picFinal.Width, vbPixels, vbTwips) + (Me.Width - Me.ScaleX(Me.ScaleWidth, vbPixels, vbTwips))
    Me.Height = Me.ScaleY(picFinal.Height, vbPixels, vbTwips) + (Me.Height - Me.ScaleY(Me.ScaleHeight, vbPixels, vbTwips))
    
    ' Reposition Signature
    lblAuthor.Top = picFinal.Height - lblAuthor.Height
    lblAuthor.Left = 0
    lblURL.Top = picFinal.Height - lblURL.Height
    lblURL.Left = picFinal.Width - lblURL.Width
    
    ' Tile background of final image with base image
    For x = 0 To picFinal.Width Step picBase.Width
        For y = 0 To picFinal.Height Step picBase.Height
            picFinal.PaintPicture picBase, x, y
        Next
    Next
    
    picBase.AutoSize = False
    picOverlay.AutoSize = False
    
    ' Resize and Tile background of base
    w = picBase.Width
    h = picBase.Height
    picBase.Width = picFinal.Width
    picBase.Height = picFinal.Height
    
    For x = 0 To picFinal.Width Step w
        For y = 0 To picFinal.Height Step h
            picBase.PaintPicture picBase, x, y
        Next
    Next
    
    ' Resize and tile background of overlay
    w = picOverlay.Width
    h = picOverlay.Height
    picOverlay.Width = picFinal.Width
    picOverlay.Height = picFinal.Height
    
    For x = 0 To picFinal.Width Step w
        For y = 0 To picFinal.Height Step h
            picOverlay.PaintPicture picOverlay, x, y
        Next
    Next
    
    ' Loop through each pixel
    For x = 0 To picFinal.Width
        For y = 0 To picFinal.Height
        
            ' Get each pictures color values at the specified position
            baseRGB = picBase.Point(x, y)
            alphaRGB = picAlpha.Point(x, y)
            overlayRGB = picOverlay.Point(x, y)
            
            ' Compute the blended pixel
            newRGB = Blend(baseRGB, alphaRGB, overlayRGB)
            
            ' paint the final image with the pixel
            picFinal.PSet (x, y), newRGB
            
        Next
        
        ' If user clicked on picture, then reset and do a different combination
        If DoAnother Then
            DoAnother = False
            picBase.AutoSize = True
            picOverlay.AutoSize = True
            picAlpha.AutoRedraw = True
            Timer1.Enabled = True
            Exit Sub
        End If
    
        ' As if it were a progress bar, display the pixels in the vertical column
        ' that we just processed
        DoEvents
    
    Next
    
    ' Save picture
    SavePicture picFinal.Image, App.Path & "\Final.bmp"
    
End Sub
