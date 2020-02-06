Attribute VB_Name = "DeclareFunctions"
Option Explicit

' キーボード関連
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long


' マウス関連
Public Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" ( _
            ByVal dwFlags As Long, _
            Optional ByVal dx As Long = 0, _
            Optional ByVal dy As Long = 0, _
            Optional ByVal dwDate As Long = 0, _
            Optional ByVal dwExtraInfo As Long = 0)

' その他
Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbsize As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
