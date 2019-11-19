Option Compare Database
Option Explicit

Private Const CurrentModName = "modMousePointers"

Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Function MouseCursor(CursorType As Long)

' Example:  =MouseCursor(32512)     ' using Public Constants from above

  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function
