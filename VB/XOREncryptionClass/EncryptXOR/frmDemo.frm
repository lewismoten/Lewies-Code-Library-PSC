VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XOR Encryption Demonstration"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdHexDecrypt 
      Caption         =   "HexDecrypt"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHexEncrypt 
      Caption         =   "HexEncrypt"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Key:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDemo.frx":0000
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   5175
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EncryptXOR As clsEncryptXOR

Private Sub cmdDecrypt_Click()
    txtMessage.Text = EncryptXOR.Decrypt(txtMessage.Text)
End Sub

Private Sub cmdEncrypt_Click()
    txtMessage.Text = EncryptXOR.Encrypt(txtMessage.Text)
End Sub

Private Sub cmdHexDecrypt_Click()
    txtMessage.Text = EncryptXOR.HexDecrypt(txtMessage.Text)
End Sub

Private Sub cmdHexEncrypt_Click()
    txtMessage.Text = EncryptXOR.HexEncrypt(txtMessage.Text)
End Sub

Private Sub Form_Load()
    Set EncryptXOR = New clsEncryptXOR
    txtKey.Text = EncryptXOR.Key
End Sub

Private Sub txtKey_Change()
    EncryptXOR.Key = txtKey.Text
End Sub
