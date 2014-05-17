VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Настройки подключения к MySQL серверу"
   ClientHeight    =   4275
   ClientLeft      =   6435
   ClientTop       =   480
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrinter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtDatabase 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CheckBox chkAOT 
      Caption         =   "Отображать окно  на верху"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Выход"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame frmScreen 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
      Begin VB.OptionButton optCorner 
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Bottom Right"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optCorner 
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         ToolTipText     =   "Top Right"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optCorner 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Bottom Left"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optCorner 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Top Left"
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Имя принтера:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   495
      TabIndex        =   20
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Пароль:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1095
      TabIndex        =   18
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Имя базы данных:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Логин:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1170
      TabIndex        =   13
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Порт:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1260
      TabIndex        =   10
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IP адресс сервера:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   -360
      TabIndex        =   9
      Top             =   120
      Width           =   2100
   End
   Begin VB.Label lblCrnr 
      AutoSize        =   -1  'True
      Caption         =   "Угол активации:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1395
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iCorner As Integer
Private bAOT As Boolean

Private Sub chkAOT_Click()
    If chkAOT.Value = 0 Then bAOT = False Else bAOT = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    WriteKey_IPAddress Me.txtIP
    WriteKey_ServerPort Me.txtPort
    WriteKey_Login Me.txtLogin
    WriteKey_Password Me.TxtPassword
    WriteKey_Database Me.txtDatabase
    WriteKey_Printer Me.txtPrinter
    Call LoaddAllSetting
    Corner = iCorner
    AOT = bAOT
    frmDisplay.UpdCorners
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtIP = IPAddress()
    Me.txtPort = ServerPort()
    Me.txtLogin = Login()
    Me.TxtPassword = Password()
    Me.txtDatabase = Database()
    Me.txtPrinter = LoadPrinter()
    frmDisplay.tmr.Enabled = False
    optCorner(Corner).Value = True
    If AOT Then chkAOT.Value = 1 Else chkAOT.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDisplay.tmr.Enabled = True
    frmDisplay.SwimUp
End Sub


Private Sub optCorner_Click(Index As Integer)
    iCorner = Index
End Sub

