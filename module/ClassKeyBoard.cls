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
 
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const INPUT_KEYBOARD As Integer = 1
 
' dwFlags
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP       As Long = &H2

Dim inputs() As INPUT_TYPE
Dim specialKey As Long
Dim specialKeyFlag As Boolean
Dim specialKeyNext As Integer

' 全体で24Byteとするように Long型（4Byte）を追加している。
Private Type KEYBD_INPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
    no_use1 As Long
    no_use2 As Long
End Type

Private Type INPUT_TYPE
    dwType As Long
    ki As KEYBD_INPUT
End Type


Public Function Send() As Long
    Send = SendInput(UBound(inputs), inputs(0), LenB(inputs(0)))
    
    ' inputs初期化
    ReDim inputs(0)
    specialKeyFlag = False
    specialKeyNext = 0
End Function

Public Sub AddKey(key As Long)
    ReDim Preserve inputs(UBound(inputs) + 1)
   
    CreateKey inputs(UBound(inputs) - 1), key, IIf(specialKeyNext <> 2, 0, 2)
    
    'Alt, Ctrl, Shiftキーのいずれかの時
    If specialKeyFlag = True Then
        specialKeyNext = specialKeyNext + 1
        If specialKeyNext = 2 Then
            AddKey specialKey
        ElseIf specialKeyNext = 3 Then
            specialKeyNext = 0
            specialKey = 0
            specialKeyFlag = False
        End If
    ElseIf VK_SHIFT_UP <= key And key <= VK_MENU_UP Then
        specialKey = key
        specialKeyFlag = True
        specialKeyNext = 1
    End If
End Sub

Private Sub CreateKey(inputs As INPUT_TYPE, ByVal key As Long, status As Long)
    inputs.dwType = INPUT_KEYBOARD
    With inputs.ki
        .wVk = key
        .wScan = MapVirtualKey(key, 0)
        .dwFlags = IIf(status = 2, KEYEVENTF_KEYUP, 0)
        .time = 0
        .dwExtraInfo = 0
        .no_use1 = 0
        .no_use2 = 0
    End With
End Sub

Private Sub Class_Initialize()
    ReDim inputs(0)
    specialKeyFlag = False
    specialKeyNext = 0
End Sub

