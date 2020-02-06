VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' キーコード一覧
' @see http://kts.sakaiweb.com/virtualkeycodes.html
'

Option Explicit
 
Private Const INPUT_KEYBOARD As Integer = 1
 
' dwFlags
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP       As Long = &H2

Dim inputs() As INPUT_TYPE

Public Function Send() As Long
    Send = SendInput(UBound(inputs), inputs(0), LenB(inputs(0)))
    
    ' inputs初期化
    ReDim inputs(0)
End Function

Public Sub AddKey(key As Long)
    Dim vlSpecialKey
    ReDim Preserve inputs(UBound(inputs) + 1)
    
    
    vlSpecialKey = key And &HFF
    
    'Alt, Ctrl, Shiftキーのいずれかの時
    If vlSpecialKey <= VK_MENU_UP And vlSpecialKey >= VK_SHIFT_UP Then

        CreateKey inputs(UBound(inputs) - 1), vlSpecialKey, IIf((key And &HF00) > 0, 1, 2)
    Else
        CreateKey inputs(UBound(inputs) - 1), key, 0
    End If
End Sub

Private Sub CreateKey(inputs As INPUT_TYPE, ByVal key As Long, status As Long)
    Dim ki As KEYBD_INPUT
    Dim vSpKey As Long
    
    inputs.dwType = INPUT_KEYBOARD
    With ki
        .wVk = key
        .wScan = MapVirtualKey(key, 0)
        .dwFlags = IIf(status = 2, KEYEVENTF_KEYUP, 0)
        .time = 0
        .dwExtraInfo = 0
        .no_use1 = 0
        .no_use2 = 0
    End With
    
    CopyMemory inputs.inputs(0), ki, LenB(ki)
End Sub

Private Sub Class_Initialize()
    ReDim inputs(0)
End Sub
