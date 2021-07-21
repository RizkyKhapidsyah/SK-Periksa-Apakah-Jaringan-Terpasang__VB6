Attribute VB_Name = "Mesin"
Option Explicit

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_NETWORK = 63
   
' Returns True if a Network is found (read-only)
Public Function ApakahJaringanTerpasang() As Boolean
   ApakahJaringanTerpasang = GetSystemMetrics(SM_NETWORK)
End Function

