VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Binary Language Utility"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFromBinary 
      Caption         =   "<-- &String"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtBinary 
      Height          =   1455
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton cmdToBinary 
      Caption         =   "&Binary -->"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtString 
      Height          =   1485
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "01"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "http://www.lewis.moten.com"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1620
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFromBinary_Click()
    Dim liCount As Long
    Dim liValue As Integer
    Dim lsResult As String
    Dim lvArray(7) As Integer
    For liCount = 1 To Len(txtBinary) Step 8
        For liCount2 = 0 To 7
            lvArray(liCount2) = Mid(txtBinary, liCount + liCount2, 1)
        Next
        lsResult = lsResult & RevertBinary(lvArray)
    Next
    txtString = lsResult
End Sub

Private Sub cmdToBinary_Click()
    Dim liCount As Long
    Dim liValue As Integer
    Dim lsResult As String
    For liCount = 1 To Len(txtString)
        liValue = Asc(Mid(txtString, liCount, 1))
        lsResult = lsResult & ConvertBinary(liValue)
    Next
    txtBinary = lsResult
End Sub

Function ConvertBinary(ByVal SourceNumber As Integer)
    Dim lsResult As String
    Dim llTemp As Long
    Dim liCount As Integer

    For liCount = 7 To 0 Step -1
        llTemp = Int(SourceNumber / (2 ^ liCount))
        lsResult = lsResult & CStr(llTemp)
        SourceNumber = SourceNumber - (llTemp * (2 ^ liCount))
    Next
    ConvertBinary = lsResult
End Function
Function RevertBinary(ByRef SourceArray() As Integer)
    For int_Count = 0 To 7 Step 1
        If int_Count < 7 Then
            SourceArray(int_Count) = SourceArray(int_Count) * (2 ^ (7 - int_Count))
        End If
        Results = Results + SourceArray(int_Count)
    Next
    RevertBinary = Chr(Results)
End Function

