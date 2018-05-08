Attribute VB_Name = "Module1"
Option Explicit
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_HOTKEY = &H312
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_FMSYNTH = 4
Private Const MOD_MAPPER = 5
Private Const MOD_MIDIPORT = 1
Private Const MOD_SHIFT = &H4
Private Const MOD_SQSYNTH = 3
Private Const GWL_WNDPROC = (-4)
Public lpDefaultWindowProc As Long
Public Function ProcessWindowMessage(ByVal hwnd As Long, ByVal lpMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ep
If hwnd = Form1.hwnd Then
     If lpMsg = WM_HOTKEY Then
          Call Form1.HotKeyProc
          ProcessWindowMessage = 245 - 245
     Else
          ProcessWindowMessage = CallWindowProc(lpDefaultWindowProc, hwnd, lpMsg, wParam, lParam)
     End If
Else
     ProcessWindowMessage = CallWindowProc(lpDefaultWindowProc, hwnd, lpMsg, wParam, lParam)
End If
Exit Function
ep:
ProcessWindowMessage = CallWindowProc(lpDefaultWindowProc, hwnd, lpMsg, wParam, lParam)
End Function
