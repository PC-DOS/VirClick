VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   300
      Left            =   5325
      TabIndex        =   5
      Top             =   3300
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":000C
      Top             =   2730
      Width           =   5340
   End
   Begin VB.Image Image3 
      Height          =   75
      Index           =   1
      Left            =   0
      Picture         =   "frmAbout.frx":0056
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   6600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   6360
      Top             =   0
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmAbout.frx":04C2
      Top             =   2085
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基於Microsoft(R) Visual Studio(R) 6.00 製作"
      Height          =   180
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   2370
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC-DOS Workshop開發"
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   1995
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本1.00"
      Height          =   180
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   1635
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Mouse Click"
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   1275
      Width           =   1710
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   0
      Left            =   210
      Picture         =   "frmAbout.frx":0D8C
      Top             =   1290
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   0
      Picture         =   "frmAbout.frx":1416
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   6480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Sub RunTestDemo()
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
On Error Resume Next
ShowCursor False
End Sub
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim l As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle l
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For l = Len(szExeName) To 1 Step -1
If Mid$(szExeName, l, 1) = "\" Then
Exit For
End If
Next l
szPathName = Left$(szExeName, l)
Exit Sub
End If
Loop Until (Process32Next(l, my) < 1)
End If
CloseHandle l
End If
End Sub
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Sub Command1_Click()
On Error Resume Next
With Form1.HotKeyGetter
.Enabled = True
.Interval = 100
End With
With Form1.MousePosGetter
.Enabled = True
.Interval = 100
End With
With Form1.VClick
.Enabled = False
End With
ShowCursor True
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Form1.Check2.Value = 1 Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If Form1.Check2.Value = 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
Unload Me
On Error Resume Next
Dim rtn As Long
If Form1.Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 192, LWA_ALPHA
End If
If Form1.Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End If
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
With Form1.HotKeyGetter
.Enabled = True
.Interval = 100
End With
With Form1.MousePosGetter
.Enabled = True
.Interval = 100
End With
With Form1.VClick
.Enabled = False
End With
ShowCursor True
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Form1.Check2.Value = 1 Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
If Form1.Check2.Value = 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
On Error Resume Next
Dim rtn As Long
If Form1.Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 192, LWA_ALPHA
End If
If Form1.Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End If
End Sub
Private Sub Form_Terminate()
On Error Resume Next
With Form1.HotKeyGetter
.Enabled = True
.Interval = 100
End With
With Form1.MousePosGetter
.Enabled = True
.Interval = 100
End With
With Form1.VClick
.Enabled = False
End With
ShowCursor True
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Form1.Check2.Value = 1 Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
If Form1.Check2.Value = 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
On Error Resume Next
Dim rtn As Long
If Form1.Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 192, LWA_ALPHA
End If
If Form1.Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
With Form1.HotKeyGetter
.Enabled = True
.Interval = 100
End With
With Form1.MousePosGetter
.Enabled = True
.Interval = 100
End With
With Form1.VClick
.Enabled = False
End With
ShowCursor True
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Form1.Check2.Value = 1 Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
If Form1.Check2.Value = 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
On Error Resume Next
Dim rtn As Long
If Form1.Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 192, LWA_ALPHA
End If
If Form1.Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Form1.hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
End If
End Sub
