VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   Caption         =   "Печать пречеков"
   ClientHeight    =   4020
   ClientLeft      =   8715
   ClientTop       =   900
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6555
   Begin VB.CommandButton cmdClear 
      Caption         =   "Очистить"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Timer tmr 
      Interval        =   5000
      Left            =   3240
      Top             =   3480
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Файл"
      WindowList      =   -1  'True
      Begin VB.Menu mnuOptions 
         Caption         =   "&Опции"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Показать"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Спрятать"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnload 
         Caption         =   "&Выход"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private R As Long
Private OldX As Single
Private OldY As Single
Private TpX As Integer ' TwipsPerPixelX
Private HotX As Integer
Private HotY As Integer

Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub


Private Sub Form_Load()
TpX = Screen.TwipsPerPixelX
    LoadInits
    Call LoaddAllSetting
    UpdCorners
    With NID
        .hIcon = Me.Icon.Handle
        .hWnd = Me.hWnd
        .szTip = "Amici precheck printer" & Chr$(0)
        .uFlags = NIF_ICON + NIF_MESSAGE + NIF_TIP
        .cbSize = Len(NID)
        NID.uCallbackMessage = WM_LBUTTONDOWN
    End With
    
    ' Adding Icon to System Tray
    Shell_NotifyIcon NIM_ADD, NID
    ' Making Window Topmost
    SwimUp
    Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static bPressed As Boolean
If Y = 0 Then ' SysTray Icon Events
    Select Case X
        Case 512 * TpX ' MouseMove
        Case 513 * TpX ' LeftButtonDown
        Case 514 * TpX ' LeftButtonUp
        Case 515 * TpX ' LeftDblClick
            Show
        Case 516 * TpX ' RightButtonDown
            bPressed = True
        Case 517 * TpX ' RightButtonUp
            If bPressed Then
                Me.PopupMenu mnuPopUp
                bPressed = False
            End If
        Case 518 * TpX ' RightButtonDblClick
    End Select
    Exit Sub
End If

If Button = 2 Then
    SwimDown
    Me.PopupMenu mnuPopUp
Else
    MousePointer = vbSizeAll
    OldX = X
    OldY = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SwimUp
    
    If Button = 1 Then
        MousePointer = vbSizeAll
        Move Left + X - OldX, Top + Y - OldY
    Else
        MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, NID
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    SaveSetting "GetWinClass", "Options", "AOT", AOT
    SaveSetting "GetWinClass", "Options", "Corner", Corner
End Sub

Private Sub mnuHide_Click()
    Hide
End Sub

Private Sub mnuOptions_Click()
    SwimDown
    frmOptions.Show
    'vbModal
    SwimUp
End Sub

Private Sub mnuShow_Click()
    SwimUp
    Show
End Sub

Private Sub mnuUnload_Click()
    If MsgBox("Выйти с программы ?", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then Unload Me
End Sub

Sub tmr_Timer()
    On Error GoTo err
    tmr.Enabled = False
    If ExistTaskPrint() = True Then
        Pprint "--------------------------------"
        Pprint "Задания печати выполнены"
        Pprint "--------------------------------"
        'DisplayPrinter
        'Pprint "--------------------------------"
    End If
    tmr.Enabled = True
    Exit Sub
err:
    Pprint "Ошибка 01: " & err.Description
    err.Clear
End Sub

Public Sub Pprint(Str$)
    If Len(Text1.Text) > 5000 Then Text1.Text = ""
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = Str$ & vbCrLf
    logs Str
    DoEvents
End Sub

Private Sub Bold(bBold As Boolean)
    FontBold = bBold
End Sub

Public Sub SwimUp()
    If AOT Then SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub SwimDown()
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub LoadInits()
    AOT = GetSetting("GetWinClass", "Options", "AOT", True)
    Corner = Val(GetSetting("GetWinClass", "Options", "Corner", 0))
End Sub

Public Sub UpdCorners()
With Screen
    Select Case Corner
        Case 0
            HotX = 0
            HotY = 0
        Case 1
            HotX = .Width \ .TwipsPerPixelX - 1
            HotY = 0
        Case 2
            HotX = 0
            HotY = .Height \ .TwipsPerPixelY - 1
        Case 3
            HotX = .Width \ .TwipsPerPixelX - 1
            HotY = .Height \ .TwipsPerPixelY - 1
    End Select
End With
End Sub

Private Sub Form_Resize()
Dim iw&, ih&
    On Error GoTo err
    If Me.Width <= 6675 Then Me.Width = 6675
    If Me.Height <= 4755 Then Me.Height = 4755
    Me.Text1.Width = 6375 + Me.Width - 6675
    Me.Text1.Height = 3255 + Me.Height - 4755
    Me.cmdClear.Left = 3720 + Me.Width - 6675
    Me.cmdClear.Top = 3480 + Me.Height - 4755
    Me.cmdOK.Left = 5160 + Me.Width - 6675
    Me.cmdOK.Top = 3480 + Me.Height - 4755
    Exit Sub
err:
End Sub

Public Sub logs(Str As String)
    Dim fTxtFileName$
    On Error GoTo err
    fTxtFileName = App.Path & "\LOG\"
    If Dir(fTxtFileName, vbDirectory) = "" Then
        MkDir fTxtFileName
    End If
    fTxtFileName = fTxtFileName & Format(Date, "DD-MM-YYYY") & ".log"
    Open fTxtFileName For Append As #1
    Print #1, Str
    Close #1
err:
End Sub

