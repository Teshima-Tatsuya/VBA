VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
' SendInput for Mouse
Private Const MOUSEEVENTF_MOVE       As Long = &H1 'マウスを移動する
Private Const MOUSEEVENTF_ABSOLUTE   As Long = &H8000& '絶対座標指定
Private Const MOUSEEVENTF_LEFTDOWN   As Long = &H2 '左 ボタンを押す
Private Const MOUSEEVENTF_LEFTUP     As Long = &H4 '左 ボタンを離す
Private Const MOUSEEVENTF_RIGHTDOWN  As Long = &H8 '右 ボタンを押す
Private Const MOUSEEVENTF_RIGHTUP    As Long = &H10 '右 ボタンを離す
Private Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20 '中央 ボタンを押す
Private Const MOUSEEVENTF_MIDDLEUP   As Long = &H40 '中央 ボタンを押す
Private Const MOUSEEVENTF_WHEEL      As Long = &H800 'ホイールを回転する
 
Private Const INPUT_MOUSE    As Integer = 0
Dim inputs() As INPUT_TYPE

Public Sub Move(p As Point)
    Dim scX As Long
    Dim scY As Long
    
    scX = 65535 / GetSystemMetrics(SM_CXSCREEN)
    scY = 65535 / GetSystemMetrics(SM_CYSCREEN)

    Add MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, p.x * scX, p.y * scY
    Send
End Sub

Public Sub RightClick()
    Add MOUSEEVENTF_RIGHTDOWN
    Add MOUSEEVENTF_RIGHTUP
    Send
End Sub

Public Sub LeftClick()
    Add MOUSEEVENTF_LEFTDOWN
    Add MOUSEEVENTF_LEFTUP
    Send
End Sub

Public Sub LeftDoubleClick()
    LeftClick
    LeftClick
End Sub

Public Sub LeftDrag(dst_p As Point)
    Dim scX As Long
    Dim scY As Long
    
    scX = 65535 / GetSystemMetrics(SM_CXSCREEN)
    scY = 65535 / GetSystemMetrics(SM_CYSCREEN)
    
    Add MOUSEEVENTF_LEFTDOWN
    Add MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, dst_p.x * scX, dst_p.y * scY
    Add MOUSEEVENTF_LEFTUP
    Send
End Sub


Public Sub Send()
    Dim i As Integer
    
    For i = 0 To UBound(inputs)
        Call SendInput(1, inputs(i), LenB(inputs(i)))
        Sleep (200)
    Next
    
    ReDim inputs(0)
End Sub

Public Sub Add(dwFlags, Optional dx As Long = 0, Optional dy As Long = 0)
    ReDim Preserve inputs(UBound(inputs) + 1)
    
    CreateMouseInput inputs(UBound(inputs) - 1), dwFlags, dx, dy
End Sub

Private Sub CreateMouseInput(inputs As INPUT_TYPE, dwFlags, Optional dx As Long = 0, Optional dy As Long = 0)
    Dim mi As MOUSE_INPUT
    inputs.dwType = INPUT_MOUSE
    
    With mi
        .dx = dx
        .dy = dy
        .dwFlags = dwFlags
        .mouseData = 0
        .time = 0
        .dwExtraInfo = 0
    End With
    
    CopyMemory inputs.inputs(0), mi, LenB(mi)
End Sub

Private Sub Class_Initialize()
    ReDim inputs(0)
End Sub
