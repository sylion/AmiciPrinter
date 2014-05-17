Attribute VB_Name = "Start_only"
Option Explicit

Public Declare Function GetWindowText Lib "user32" _
        Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
       
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
        
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type WINDOWPLACEMENT
  Length As Long
  FLAGS As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Dim strCaptions() As String
Dim lngHandle() As Long

Sub Main()
  If RestorePrevInstance("«агрузка продаж в SQL сервер") Then
    Exit Sub
  Else
    frmDisplay.Show
    frmDisplay.tmr.Enabled = False
    frmDisplay.tmr.Enabled = True
    frmDisplay.Hide
  End If
End Sub

Public Function RestorePrevInstance(strCaption As String) As Boolean
    Dim iCount As Integer
    Dim i As Long
    Dim Pos As Integer
    Dim lngEnum As Long
    Dim udtCurrWin As WINDOWPLACEMENT
    Dim lngLenArray As Long
    
    ReDim strCaptions(0)
    ReDim lngHandle(0) ' то же чистим
    lngEnum = EnumWindows(AddressOf Callback1_EnumWindows, 0)
    iCount = 0
    lngLenArray = UBound(strCaptions)
    For i = 0 To lngLenArray
        Pos = InStr(1, strCaptions(i), strCaption, vbTextCompare)
        If Pos > 0 Then
            udtCurrWin.Length = Len(udtCurrWin)
            Call GetWindowPlacement(lngHandle(i), udtCurrWin)
            If udtCurrWin.showCmd = SW_SHOWMINIMIZED Then
                udtCurrWin.Length = Len(udtCurrWin)
                udtCurrWin.FLAGS = 0&
                udtCurrWin.showCmd = SW_SHOWNORMAL
                Call SetWindowPlacement(lngHandle(i), udtCurrWin)
            End If
            Call SetForegroundWindow(lngHandle(i))
            iCount = iCount + 1
        End If
    Next
    
    If iCount >= 1 Then
        RestorePrevInstance = True
    Else
        RestorePrevInstance = False
    End If
  
End Function

Public Function Callback1_EnumWindows(ByVal hWnd As Long, ByVal lpData As Long) As Long
    Dim cnt As Long
    Dim strTitle As String * 256
    
    cnt = GetWindowText(hWnd, strTitle, 255)
    
    If cnt > 0 Then
        ReDim Preserve lngHandle(UBound(strCaptions) + 1)
        ReDim Preserve strCaptions(UBound(strCaptions) + 1)
        strCaptions(UBound(strCaptions)) = Left$(strTitle, cnt)
        lngHandle(UBound(lngHandle)) = hWnd
    End If
    
    Callback1_EnumWindows = 1
End Function

