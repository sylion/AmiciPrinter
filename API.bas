Attribute VB_Name = "API"
Option Explicit

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const SW_NORMAL = 1

Public Type POINTAPI
   X As Long
   Y As Long
End Type
 
Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Uf
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
 
Public Enum dwMess
    NIM_ADD = &H0
    NIM_DELETE = &H2
    NIM_MODIFY = &H1
End Enum

Public Enum Uf
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
End Enum

Public Enum CallMess
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    WM_SETFOCUS = &H7
    WM_KEYDOWN = &H100
    WM_KEYFIRST = &H100
    WM_KEYLAST = &H108
    WM_KEYUP = &H101
End Enum
 
 
Public Declare Function GetCursorPos _
Lib "user32" ( _
    lpPoint As POINTAPI _
) As Long
      
Public Declare Function WindowFromPointXY _
Lib "user32" Alias "WindowFromPoint" ( _
    ByVal xPoint As Long, _
    ByVal yPoint As Long _
) As Long
      
Public Declare Function GetModuleFileName _
Lib "kernel32" Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long _
) As Long
      
Public Declare Function GetWindowWord _
Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Integer
      
Public Declare Function GetWindowLong _
Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long
      
      
Public Declare Function GetParent _
Lib "user32" ( _
    ByVal hWnd As Long _
) As Long
      
      
Public Declare Function GetClassName _
Lib "user32" Alias "GetClassNameA" ( _
    ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long
  
Public Declare Function GetWindowText _
Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long
   
Public Declare Function SetWindowPos _
Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long
   
Public Declare Function Shell_NotifyIcon _
Lib "shell32.dll" Alias "Shell_NotifyIconA" ( _
    ByVal dwMessage As dwMess, _
    lpData As NOTIFYICONDATA _
) As Long
   
Public Declare Function SetCapture _
Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Public Declare Function ReleaseCapture _
Lib "user32" () As Long

Public Declare Function ShellExecute _
Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Public Declare Sub Sleep _
Lib "kernel32" (ByVal dwMilliseconds As Long)
      
Public NID As NOTIFYICONDATA
Public Corner As Integer
Public AOT As Boolean
