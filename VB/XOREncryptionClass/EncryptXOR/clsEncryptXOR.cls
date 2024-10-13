VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryptXOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Author:
'   Lewis Edward Moten III
'   lewis@moten.com
'   http://www.lewismoten.com

' Created:
'   March 2, 2002 10:11 AM GMT -5 EST

' Title:
'   XOR Encryption Class

' Description:
'   Encrypts and Decrypts messages using the XOR algorithm
'   with support of unicode & method of getting pas the NULL
'   barrier by converting messages to hex strings.
'   Function of encryption are exposed through a Class Module.
'   Create a new class called clsEncryptXOR and paste this code
'   into it.

' Comments:
'   XOR is an advanced algorithm based off of the Enigma machine
'   used in World War II. XOR is not secure.  It has been proven
'   that XOR encryptions can be easily broken in a matter of seconds.
'   To boost security of your XOR encryption, it helps to have longer
'   Keys.  Even better to have a key that is longer then your message.

' Functions:
'   HexEncrypt
'       Expects that you will pass a normal message to it and expect a hex string in return.
'   HexDecrypt
'       Expects that you pass a hex string to the procedure and expect a normal message in return.
'   Encrypt
'       Pass procedure a normal string and get the encrypted string back
'   Decrypt
'       Pass procedure an encrypted string and get decrypted string back

' Properties:
'   Key
'       password to encrypt/decrypt with.
'       Longer keys ensure better security

' Demo:
'   Dim EncryptXOR As clsEncryptXOR
'   Set EncryptXOR = New clsEncryptXOR
'   EncryptXOR.Key = "password"
'   MsgBox EncryptXOR.Decrypt(EncryptXOR.Encrypt("Hello World"))
'   MsgBox EncryptXOR.HexDecrypt(EncryptXOR.HexEncrypt("Hello World"))
'   Set EncryptXOR = Nothing

Public Key As String

Private Sub Class_Initialize()
    Key = "password"
End Sub

Private Function ToHex(ByRef pstrMessage As String) As String
    
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrHex As String
    
    llngMaxIndex = LenB(pstrMessage)
    
    For llngIndex = 1 To llngMaxIndex
        lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(pstrMessage, llngIndex, 1))), 2)
    Next
    
    ToHex = lstrHex
    
End Function

Private Function FromHex(ByRef pstrHex As String) As String
    
    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrMessage As String
    
    llngMaxIndex = Len(pstrHex)
    
    For llngIndex = 1 To llngMaxIndex Step 2
        lstrMessage = lstrMessage & ChrB("&h" & Mid(pstrHex, llngIndex, 2))
    Next
    
    FromHex = lstrMessage
    
End Function

Public Function Decrypt(ByVal pstrMessage As String) As String
    pstrMessage = ToHex(pstrMessage)
    pstrMessage = Translate(pstrMessage)
    pstrMessage = FromHex(pstrMessage)
    Decrypt = pstrMessage
End Function
Public Function Encrypt(ByVal pstrMessage As String) As String
    pstrMessage = ToHex(pstrMessage)
    pstrMessage = Translate(pstrMessage)
    pstrMessage = FromHex(pstrMessage)
    Encrypt = pstrMessage
End Function
Public Function HexDecrypt(ByVal pstrMessage As String) As String
    pstrMessage = Translate(pstrMessage)
    pstrMessage = FromHex(pstrMessage)
    HexDecrypt = pstrMessage
End Function
Public Function HexEncrypt(ByVal pstrMessage As String) As String
    pstrMessage = ToHex(pstrMessage)
    pstrMessage = Translate(pstrMessage)
    HexEncrypt = pstrMessage
End Function
Private Function Translate(ByRef pstrMessage As String) As String
    
    Dim llngIndex As Long
    Dim llngMessageLength As Long
    Dim llngKeyLength As Long
    Dim lstrKey As String
    Dim llngKeyIndex As Long
    Dim lbytMessageByte As Byte
    Dim lbytKeyByte As Byte
    Dim llngMessageIndex As Long
    Dim lstrTranslation As String
    
    lstrKey = ToHex(Key)
    llngKeyLength = Len(lstrKey) \ 2
    
    If llngKeyLength = 0 Then Exit Function
    
    llngMessageLength = Len(pstrMessage) \ 2
    
    For llngIndex = 1 To llngMessageLength
        
        llngKeyIndex = (((llngIndex - 1) Mod llngKeyLength) * 2) + 1
        llngMessageIndex = ((llngIndex - 1) * 2) + 1
        
        lbytMessageByte = Int("&h" & Mid(pstrMessage, llngMessageIndex, 2))
        lbytKeyByte = Int("&h" & Mid(lstrKey, llngKeyIndex, 2))
        
        lstrTranslation = lstrTranslation & ToHex(ChrB(lbytMessageByte Xor lbytKeyByte))
        
    Next
    
    Translate = lstrTranslation
    
End Function

