Attribute VB_Name = "Module1"
Option Explicit
'Public Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hwnd As Long, ByVal hAccTable As Long, lpMsg As MSG) As Long
'Use that for menu keys but i have no idea how to do it in vb, only C
Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type MSG
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long    'You only need the 4 above, you can delete this if you want
  pt As POINTAPI
End Type

Sub DoMyEvents(ByVal hwnd As Long)
  Dim i As Long
  Dim msg1 As MSG
  i = PeekMessage(msg1, hwnd, 0, 0, 1)
  If i <> 0 Then
    Call TranslateMessage(msg1) 'For keys
    Call DispatchMessage(msg1)  'For normal Events
  End If
End Sub
