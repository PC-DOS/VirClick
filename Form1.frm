VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Virtual Mouse Click - PC-DOS Workshop"
   ClientHeight    =   4665
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6660
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   5145
      TabIndex        =   26
      Top             =   4305
      Width           =   1485
   End
   Begin VB.CheckBox Check3 
      Caption         =   "軟件窗口75%透明(&N)"
      Height          =   285
      Left            =   2370
      TabIndex        =   25
      Top             =   4305
      Width           =   2085
   End
   Begin VB.CheckBox Check2 
      Caption         =   "軟件窗口置於頂層(&P)"
      Height          =   285
      Left            =   45
      TabIndex        =   24
      Top             =   4305
      Width           =   2175
   End
   Begin 工程1.cSysTray cSysTray1 
      Left            =   5385
      Top             =   3450
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":068A
      TrayTip         =   "Virtual Mouse Cilck - 正在執行操作"
   End
   Begin VB.Frame Frame5 
      Caption         =   "當前鼠標位置"
      Height          =   630
      Left            =   30
      TabIndex        =   19
      Top             =   3630
      Width           =   6615
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   240
         Left            =   4335
         TabIndex        =   23
         Top             =   255
         Width           =   2160
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y軸位置:"
         Height          =   180
         Left            =   3570
         TabIndex        =   22
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   240
         Left            =   990
         TabIndex        =   21
         Top             =   255
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X軸位置:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Timer VClick 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4050
      Top             =   525
   End
   Begin VB.Timer HotKeyGetter 
      Interval        =   100
      Left            =   2445
      Top             =   450
   End
   Begin VB.Timer MousePosGetter 
      Interval        =   100
      Left            =   1350
      Top             =   450
   End
   Begin VB.Frame Frame3 
      Caption         =   "程序附加選項"
      Height          =   1170
      Left            =   30
      TabIndex        =   11
      Top             =   2385
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "開始模擬操作後,不允許移動鼠標指針(&V)"
         Height          =   315
         Left            =   135
         TabIndex        =   18
         Top             =   1515
         Visible         =   0   'False
         Width           =   6360
      End
      Begin VB.Frame Frame4 
         Caption         =   "模擬操作開始後"
         Height          =   570
         Left            =   135
         TabIndex        =   14
         Top             =   510
         Width           =   6345
         Begin VB.OptionButton Option4 
            Caption         =   "不執行操作(&O)"
            Height          =   240
            Left            =   150
            TabIndex        =   17
            Top             =   255
            Width           =   1500
         End
         Begin VB.OptionButton Option5 
            Caption         =   "最小化到系統托盤(&M)"
            Height          =   240
            Left            =   1800
            TabIndex        =   16
            Top             =   255
            Width           =   2055
         End
         Begin VB.OptionButton Option6 
            Caption         =   "窗口置頂並半透明(&T)"
            Height          =   240
            Left            =   4005
            TabIndex        =   15
            Top             =   255
            Value           =   -1  'True
            Width           =   2205
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":0D24
         Left            =   1590
         List            =   "Form1.frx":0D4C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   180
         Width           =   4890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作控制快捷鍵:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "鼠標模擬選項"
      Height          =   1290
      Left            =   45
      TabIndex        =   2
      Top             =   1065
      Width           =   6600
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   2535
         Max             =   30
         Min             =   1
         SmallChange     =   5
         TabIndex        =   9
         Top             =   405
         Value           =   5
         Width           =   3990
      End
      Begin VB.Frame Frame2 
         Caption         =   "鼠標鍵動作"
         Height          =   1035
         Left            =   135
         TabIndex        =   3
         Top             =   195
         Width           =   2250
         Begin VB.OptionButton Option3 
            Caption         =   "左鼠標鍵雙擊(&D)"
            Height          =   180
            Left            =   150
            TabIndex        =   6
            Top             =   765
            Width           =   1995
         End
         Begin VB.OptionButton Option2 
            Caption         =   "右鼠標鍵單擊(&R)"
            Height          =   180
            Left            =   150
            TabIndex        =   5
            Top             =   495
            Width           =   1995
         End
         Begin VB.OptionButton Option1 
            Caption         =   "左鼠標鍵單擊(&L)"
            Height          =   180
            Left            =   150
            TabIndex        =   4
            Top             =   210
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.Label lblDly 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         Height          =   255
         Left            =   2550
         TabIndex        =   10
         Top             =   690
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "間隔單位為秒,有效範圍為1至30"
         Height          =   180
         Left            =   2535
         TabIndex        =   8
         Top             =   1020
         Width           =   2520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作執行間隔:"
         Height          =   180
         Left            =   2535
         TabIndex        =   7
         Top             =   195
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0DBF
      Height          =   585
      Index           =   1
      Left            =   870
      TabIndex        =   1
      Top             =   435
      Width           =   5700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "這個工具可以幫助您模擬鼠標按鈕的操作,包括左鍵單擊,右鍵單擊和左鍵雙擊."
      Height          =   420
      Index           =   0
      Left            =   870
      TabIndex        =   0
      Top             =   60
      Width           =   5700
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   90
      Picture         =   "Form1.frx":0E4D
      Top             =   120
      Width           =   720
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuLock 
         Caption         =   "鎖定鼠標位置(&L)..."
      End
      Begin VB.Menu mnuHide 
         Caption         =   "鼠標指針隱藏(&C)..."
      End
      Begin VB.Menu mnuRect 
         Caption         =   "鼠標區域限制(&R)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "幫助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "關於 Virtual Mouse Click(&A)..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_HOTKEY = &H312
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_FMSYNTH = 4
Private Const MOD_MAPPER = 5
Private Const MOD_MIDIPORT = 1
Private Const MOD_SHIFT = &H4
Private Const MOD_SQSYNTH = 3
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal Id As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal Id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
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
Private Const GWL_WNDPROC = (-4)
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
Private Sub Check2_Click()
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Check2.Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If Check2.Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
On Error Resume Next
Dim rtn As Long
If Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Exit Sub
End If
If Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Exit Sub
End If
End Sub
Private Sub Check3_Click()
On Error Resume Next
Dim rtn As Long
If Check3.Value = 1 Then
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Exit Sub
End If
If Check3.Value = 0 Then
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Exit Sub
End If
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Check2.Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If Check2.Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F1
Case 1
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F2
Case 2
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F3
Case 3
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F4
Case 4
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F5
Case 5
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F6
Case 6
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F7
Case 7
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F8
Case 8
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F9
Case 9
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F10
Case 10
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F11
Case 11
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F12
End Select
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F1
Case 1
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F2
Case 2
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F3
Case 3
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F4
Case 4
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F5
Case 5
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F6
Case 6
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F7
Case 7
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F8
Case 8
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F9
Case 9
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F10
Case 10
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F11
Case 11
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F12
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
UnregisterHotKey hwnd, 245
Unload Me
End
End Sub
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
Me.Show
End Sub
Private Sub Form_Activate()
On Error Resume Next
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F1
Case 1
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F2
Case 2
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F3
Case 3
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F4
Case 4
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F5
Case 5
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F6
Case 6
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F7
Case 7
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F8
Case 8
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F9
Case 9
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F10
Case 10
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F11
Case 11
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F12
End Select
ShowCursor True
End Sub
Private Sub Form_Load()
On Error Resume Next
On Error Resume Next
lpDefaultWindowProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf ProcessWindowMessage
Select Case Combo1.ListIndex
Case 0
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F1
Case 1
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F2
Case 2
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F3
Case 3
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F4
Case 4
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F5
Case 5
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F6
Case 6
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F7
Case 7
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F8
Case 8
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F9
Case 9
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F10
Case 10
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F11
Case 11
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F12
End Select
HKStateCtrl = 0
HKStateFn = 0
Me.Option1.Value = True
Option6.Value = True
With Me.Option1
.Enabled = True
End With
With Me.Option2
.Enabled = True
End With
With Me.Option3
.Enabled = True
End With
With Me.Option4
.Enabled = True
End With
With Me.Option5
.Enabled = True
End With
With Me.Option6
.Enabled = True
End With
On Error Resume Next
dwMouseFlag = 0
dwMouseFlag = ME_LBCLICK
With Me.HScroll1
.Value = 5
.Max = 30
.Min = 1
.Enabled = True
.LargeChange = 10
.SmallChange = 5
End With
With Me.HotKeyGetter
.Interval = 100
.Enabled = False
End With
With Me.MousePosGetter
.Enabled = True
.Interval = 100
End With
With Me.VClick
.Interval = 5000
.Enabled = False
End With
With Me.Combo1
.ListIndex = 0
.Enabled = True
End With
On Error Resume Next
Dim lpCPT As POINTAPI
GetCursorPos lpCPT
With lpCPT
lpX = .x
lpY = .y
End With
With Me.lblX
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(lpX)
End With
With Me.lblY
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(lpY)
End With
With Check2
.Enabled = True
.Value = 1
End With
With Check3
.Enabled = True
.Value = 0
End With
On Error Resume Next
Const HWND_NOTOPMOST = -2
If Check2.Value = 1 Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
If Check2.Value = 0 Then
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
UnregisterHotKey hwnd, 245
Unload Me
End Sub
Private Sub Form_Terminate()
On Error Resume Next
UnregisterHotKey hwnd, 245
'Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
UnregisterHotKey hwnd, 245
Unload Me
End Sub
Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
End Sub

Private Sub HotKeyGetter_Timer()
On Error Resume Next
On Error Resume Next
Select Case Combo1.ListIndex
Case 0
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F1
Case 1
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F2
Case 2
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F3
Case 3
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F4
Case 4
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F5
Case 5
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F6
Case 6
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F7
Case 7
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F8
Case 8
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F9
Case 9
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F10
Case 10
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F11
Case 11
UnregisterHotKey Me.hwnd, 245
RegisterHotKey Me.hwnd, 245, MOD_CONTROL, VK_F12
End Select
End Sub

'Private Sub HotKeyGetter_Timer()
'On Error Resume Next
'HKStateCtrl = 0
'HKStateFn = 0
'Const HWND_NOTOPMOST = -2
'Const GKS_PRESSED = 1
'Const GKS_UNPRESSED = 0
'Const CTRL = &H11
'If 1 = 245 Then
'Select Case Me.Combo1.ListIndex
'Case 0
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F1)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'Dim rtn As Long
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'Dim rtn2 As Long
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 1
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F2)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 2
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F3)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 3
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F4)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End Select
'If Combo1.ListIndex = 4 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F5)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 5 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F6)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 6 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F7)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 7 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F8)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 8 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F9)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 9 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F10)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 10 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F11)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 11 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F12)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'End If
'Select Case Me.Combo1.ListIndex
'Case 0
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F1)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 1
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F2)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 2
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F3)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'Case 3
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F4)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'Me.Show
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End Select
'If Combo1.ListIndex = 4 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F5)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 5 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F6)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 6 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F7)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 7 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F8)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame2.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 8 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F9)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 9 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F10)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 10 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F11)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'If Combo1.ListIndex = 11 Then
'HKStateCtrl = GetKeyState(VK_CONTROL)
'HKStateFn = GetKeyState(VK_F12)
'Debug.Print HKStateCtrl
'Debug.Print HKStateFn
'Debug.Print "EQV=" & HKStateCtrl + HKStateFn
'If (HKStateCtrl + HKStateFn) = -126 Or (HKStateCtrl + HKStateFn) = -128 Then
''HKStateCtrl = 0
''HKStateFn = 0
'If Me.VClick.Enabled = False Then
'Option1.Enabled = False
'Option2.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
'Option5.Enabled = False
'Option6.Enabled = False
'Combo1.Enabled = False
'Check1.Enabled = False
'HScroll1.Enabled = False
'Frame1.Enabled = False
'Frame2.Enabled = False
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = False
'lblY.Enabled = False
'Me.Label1(0).Enabled = False
'Label1(1).Enabled = False
'Me.Label2.Enabled = False
'Label3.Enabled = False
'Label4.Enabled = False
'Me.Label5.Enabled = False
'Label6.Enabled = False
'lblDly.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False
'With Me.VClick
'.Enabled = True
'End With
'HKStateCtrl = 0
'HKStateFn = 0
'If Option6.Value = True Then
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 127, LWA_ALPHA
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'End If
'If Option5.Value = True Then
'Me.Hide
'With Me.cSysTray1
'.InTray = True
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'End If
'Exit Sub
'Else
'Option1.Enabled = True
'Option2.Enabled = True
'Option3.Enabled = True
'Option4.Enabled = True
'Option5.Enabled = True
'Option6.Enabled = True
'Combo1.Enabled = True
'Check1.Enabled = True
'HScroll1.Enabled = True
'Frame1.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame2.Enabled = True
'lpX = CLng(Me.lblX.Caption)
'lpY = CLng(lblY.Caption)
'lblX.Enabled = True
'lblY.Enabled = True
'Me.Label1(0).Enabled = True
'Label1(1).Enabled = True
'Me.Label2.Enabled = True
'Label3.Enabled = True
'Label4.Enabled = True
'Me.Label5.Enabled = True
'Label6.Enabled = True
'lblDly.Enabled = True
'Frame3.Enabled = True
'Frame4.Enabled = True
'Frame5.Enabled = True
'With Me.VClick
'.Enabled = False
'End With
'On Error Resume Next
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
'With cSysTray1
'.InTray = False
'.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
'End With
'Me.Show
'Select Case Check2.Value
'Case 1
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'Case 0
'SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
'End Select
'Select Case Check3.Value
'Case 1
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 192, LWA_ALPHA
'Case 0
'rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
'rtn = rtn Or WS_EX_LAYERED
'SetWindowLong hWnd, GWL_EXSTYLE, rtn
'SetLayeredWindowAttributes hWnd, 0, 255, LWA_ALPHA
'End Select
'HKStateCtrl = 0
'HKStateFn = 0
'Exit Sub
'End If
'Else
'Exit Sub
'End If
'End If
'End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Dim uTime As Integer
Dim uTimeMil As Integer
uTime = CInt(HScroll1.Value)
With Me.lblDly
.Alignment = 2
.BorderStyle = 1
.BackColor = RGB(255, 255, 255)
.BackStyle = 0
.Caption = CStr(uTime)
End With
uTimeMil = CInt(uTime) * 1000
With Me.VClick
.Interval = uTimeMil
.Enabled = False
End With
End Sub
Private Sub mnuAbout_Click()
With Me.HotKeyGetter
.Enabled = False
.Interval = 100
End With
With Me.MousePosGetter
.Enabled = False
.Interval = 100
End With
With Me.VClick
.Enabled = False
End With
If 1 = 2 Then
With Me
.Hide
End With
End If
frmAbout.Show
End Sub
Private Sub mnuHide_Click()
With Me.HotKeyGetter
.Enabled = False
.Interval = 100
End With
With Me.MousePosGetter
.Enabled = False
.Interval = 100
End With
With Me.VClick
.Enabled = False
End With
If 1 = 2 Then
With Me
.Hide
End With
End If
Form3.Show 1
End Sub
Private Sub mnuLock_Click()
With Me.HotKeyGetter
.Enabled = False
.Interval = 100
End With
With Me.MousePosGetter
.Enabled = False
.Interval = 100
End With
With Me.VClick
.Enabled = False
End With
If 1 = 2 Then
With Me
.Hide
End With
End If
Form2.Show 1
End Sub
Private Sub mnuRect_Click()
On Error Resume Next
With Me.HotKeyGetter
.Enabled = False
.Interval = 100
End With
With Me.MousePosGetter
.Enabled = False
.Interval = 100
End With
With Me.VClick
.Enabled = False
End With
If 1 <> 2 Then
With Me
.Hide
End With
End If
Form4.Show
End Sub
Private Sub MousePosGetter_Timer()
On Error Resume Next
Dim lpCPT As POINTAPI
GetCursorPos lpCPT
With lpCPT
lpX = .x
lpY = .y
End With
With Me.lblX
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(lpX)
End With
With Me.lblY
.Alignment = 2
.BackColor = RGB(255, 255, 255)
.BorderStyle = 1
.BackStyle = 0
.Caption = CStr(lpY)
End With
End Sub
Private Sub Option1_Click()
On Error Resume Next
dwMouseFlag = 0
dwMouseFlag = ME_LBCLICK
End Sub
Private Sub Option2_Click()
On Error Resume Next
dwMouseFlag = 0
dwMouseFlag = ME_RBCLICK
End Sub
Private Sub Option3_Click()
On Error Resume Next
dwMouseFlag = 0
dwMouseFlag = ME_LBDBLCLICK
End Sub
Private Sub VClick_Timer()
On Error Resume Next
Dim dwX As Long
Dim dwY As Long
dwX = CLng(Me.lblX.Caption)
dwY = CLng(Me.lblY.Caption)
If Check1.Value = 1 Then
dwX = lpX
dwY = lpY
SetCursorPos lpX, lpY
End If
Dim dwCurPos As POINTAPI
GetCursorPos dwCurPos
With dwCurPos
dwX = .x
dwY = .y
End With
If dwMouseFlag = ME_LBCLICK Then
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
End If
If dwMouseFlag = ME_LBDBLCLICK Then
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
End If
If dwMouseFlag = ME_RBCLICK Then
mouse_event MOUSEEVENTF_RIGHTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_RIGHTUP, dwX, dwY, 0, 0
End If
End Sub
Public Sub HotKeyProc()
Const HWND_NOTOPMOST = -2
If Me.VClick.Enabled = False Then
Select Case dwMouseFlag
Case ME_LBCLICK
Option1.Value = True
Case ME_RBCLICK
Option2.Value = True
Case ME_LBDBLCLICK
Option3.Value = True
End Select
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Combo1.Enabled = False
Check1.Enabled = False
HScroll1.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
lpX = CLng(Me.lblX.Caption)
lpY = CLng(lblY.Caption)
lblX.Enabled = False
lblY.Enabled = False
Me.Label1(0).Enabled = False
Label1(1).Enabled = False
Me.Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Me.Label5.Enabled = False
Label6.Enabled = False
lblDly.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
Me.mnuAbout.Enabled = False
Me.mnuHelp.Enabled = False
Me.mnuHide.Enabled = False
Me.mnuLock.Enabled = False
Me.mnuRect.Enabled = False
Me.mnuTools.Enabled = False
lpX = CLng(Me.lblX.Caption)
lpY = CLng(lblY.Caption)
lblX.Enabled = False
lblY.Enabled = False
Me.Label1(0).Enabled = False
Label1(1).Enabled = False
Me.Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Me.Label5.Enabled = False
Label6.Enabled = False
lblDly.Enabled = False
With Me.VClick
.Enabled = True
End With
HKStateCtrl = 0
HKStateFn = 0
If Option6.Value = True Then
On Error Resume Next
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 127, LWA_ALPHA
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End If
If Option5.Value = True Then
Me.Hide
With Me.cSysTray1
.InTray = True
.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
End With
End If
Select Case dwMouseFlag
Case ME_LBCLICK
Option1.Value = True
Case ME_RBCLICK
Option2.Value = True
Case ME_LBDBLCLICK
Option3.Value = True
End Select
With Me.VClick
.Enabled = True
End With
On Error Resume Next
Dim dwX As Long
Dim dwY As Long
dwX = CLng(Me.lblX.Caption)
dwY = CLng(Me.lblY.Caption)
If Check1.Value = 1 Then
dwX = lpX
dwY = lpY
SetCursorPos lpX, lpY
End If
Dim dwCurPos As POINTAPI
GetCursorPos dwCurPos
With dwCurPos
dwX = .x
dwY = .y
End With
If dwMouseFlag = ME_LBCLICK Then
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
End If
If dwMouseFlag = ME_LBDBLCLICK Then
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, dwX, dwY, 0, 0
End If
If dwMouseFlag = ME_RBCLICK Then
mouse_event MOUSEEVENTF_RIGHTDOWN, dwX, dwY, 0, 0
mouse_event MOUSEEVENTF_RIGHTUP, dwX, dwY, 0, 0
End If
Exit Sub
Else
Select Case dwMouseFlag
Case ME_LBCLICK
Option1.Value = True
Case ME_RBCLICK
Option2.Value = True
Case ME_LBDBLCLICK
Option3.Value = True
End Select
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Combo1.Enabled = True
Check1.Enabled = True
HScroll1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Me.mnuAbout.Enabled = True
Me.mnuHelp.Enabled = True
Me.mnuHide.Enabled = True
Me.mnuLock.Enabled = True
Me.mnuRect.Enabled = True
Me.mnuTools.Enabled = True
lpX = CLng(Me.lblX.Caption)
lpY = CLng(lblY.Caption)
lblX.Enabled = True
lblY.Enabled = True
Me.Label1(0).Enabled = True
Label1(1).Enabled = True
Me.Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Me.Label5.Enabled = True
Label6.Enabled = True
lblDly.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
With Me.VClick
.Enabled = False
End With
On Error Resume Next
Dim rtn2 As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
With cSysTray1
.InTray = False
.TrayTip = "Virtual Mouse Cilck - 正在執行操作"
End With
Me.Show
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Select
Select Case Check3.Value
Case 1
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
Case 0
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
End Select
HKStateCtrl = 0
HKStateFn = 0
Select Case dwMouseFlag
Case ME_LBCLICK
Option1.Value = True
Case ME_RBCLICK
Option2.Value = True
Case ME_LBDBLCLICK
Option3.Value = True
End Select
With Me.VClick
.Enabled = False
End With
Exit Sub
End If
End Sub
