Attribute VB_Name = "Utilites"
Option Explicit
Option Compare Text
Public Const IniFile$ = "MySQL.INI"
Public Const MySQLSection$ = "MySQL"
'----------------------------------------------------------------------------
Public vIPAddress As String
Public vServerPort As Integer
Public vLogin As String
Public vPassword As String
Public vDatabase As String
Public vPrinter As String
'----------------------------------------------------------------------------
Public Function LoaddAllSetting()
    vIPAddress = IPAddress()
    vServerPort = ServerPort()
    vLogin = Login()
    vPassword = Password()
    vDatabase = Database()
    vPrinter = LoadPrinter()
End Function

Public Function LoadPrinter() As String
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    LoadPrinter = Trim(VBGetPrivateProfileString(MySQLSection, "Printer", FileName))
    Exit Function
er:
LoadPrinter = ""
End Function

Public Function WriteKey_Printer(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "Printer", KeyValue, FileName
    Exit Function
er:
End Function

'----------------------------------------------------------------------------
Public Function IPAddress() As String
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    IPAddress = Trim(VBGetPrivateProfileString(MySQLSection, "IP", FileName))
    Exit Function
er:
IPAddress = ""
End Function

Public Function WriteKey_IPAddress(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "IP", KeyValue, FileName
    Exit Function
er:
End Function

'----------------------------------------------------------------------------
Public Function ServerPort() As Integer
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    ServerPort = Val(VBGetPrivateProfileString(MySQLSection, "PORT", FileName))
    If ServerPort <= 0 Then ServerPort = 20
    Exit Function
er:
ServerPort = 20
End Function

Public Function WriteKey_ServerPort(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "PORT", KeyValue, FileName
    Exit Function
er:
End Function

'----------------------------------------------------------------------------
Public Function Login() As String
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    Login = Trim(VBGetPrivateProfileString(MySQLSection, "LOGIN", FileName))
    Exit Function
er:
Login = ""
End Function

Public Function WriteKey_Login(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "LOGIN", KeyValue, FileName
    Exit Function
er:
End Function

'----------------------------------------------------------------------------
Public Function Password() As String
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    Password = Trim(VBGetPrivateProfileString(MySQLSection, "PASSWORD", FileName))
    Exit Function
er:
Password = ""
End Function

Public Function WriteKey_Password(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "PASSWORD", KeyValue, FileName
    Exit Function
er:
End Function

'----------------------------------------------------------------------------
Public Function Database() As String
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    Database = Trim(VBGetPrivateProfileString(MySQLSection, "DATABASE", FileName))
    Exit Function
er:
Database = ""
End Function

Public Function WriteKey_Database(KeyValue As String)
    Dim FileName$
    On Error GoTo er
    FileName$ = App.Path
    If Right$(FileName$, 1) <> "\" Then FileName$ = FileName$ & "\"
    FileName$ = FileName$ & IniFile
    WriteKey MySQLSection, "DATABASE", KeyValue, FileName
    Exit Function
er:
End Function
