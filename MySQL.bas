Attribute VB_Name = "MySQL"
Option Explicit
Option Compare Text

Private Const MsgerrCon$ = "Не удалось установить соединение с базой данных eMenu Host: "
Private Const MsgerrQuery$ = "Ощибка при выполнении запроса !" & vbCrLf

Private MySQL As MYSQLDBLib.MySQL
Private CnMYSQL, RS_Status As Long
Private temp_sql As String

Public OrderIdStr As String
Public Creator_name As String
Public Table_name As String
Public CcreatedTime As String
Public CreatedDate As String
Public comment As String
Public NumGuests As String

Function ExistTaskPrint() As Boolean
    On Error GoTo err
    ExistTaskPrint = False
    Set MySQL = Nothing
    Set MySQL = New MYSQLDBLib.MySQL
    MySQL.CloseConnection
    CnMYSQL = MySQL.OpenConnection(Trim(vIPAddress), Trim(Str(vServerPort)), Trim(vLogin), Trim(vPassword), Trim(vDatabase), 0)
    If CnMYSQL < 0 Then
        frmDisplay.Pprint MsgerrCon & Trim(vIPAddress) & "!"
        GoTo err
    End If
    temp_sql = "SELECT `orderid` as `id` FROM `account_ordered_item` WHERE `isPrinted` = 0 GROUP BY `id`"
    RS_Status = MySQL.OpenRecordset("IDs", temp_sql)
    If (RS_Status = 0) Then
        frmDisplay.Pprint MsgerrQuery & "" & Trim(MySQL.LastError)
        GoTo err
    End If
    MySQL.MoveFirst "IDs"
    'Do While (MySQL.EOR("IDs") = 0)
        DoPrint (Val(MySQL.GetFieldByName("IDs", "id")))
    '    MySQL.MoveNext "IDs"
    'Loop
    'MySQL.CloseRecordset "isPrinted"
close_cnn:
    MySQL.CloseConnection
    Set MySQL = Nothing
    Exit Function
err:
    frmDisplay.Pprint "Ошибка: " & err.Description
    err.Clear
    MySQL.CloseConnection
    On Error Resume Next
    ExistTaskPrint = False
    GoTo close_cnn
End Function

Function DoPrint(ID As String) As Boolean
    Dim sql As String
    sql = "SELECT `t`.`name_ua` AS `table`, `a`.`name_ua` AS `creator`, `s`.`comment` AS `comment`, FROM_UNIXTIME(`s`.`createdtime`) AS `date`"
    sql = sql & " FROM (SELECT `creatorid`, `tableId`, `createdTime`, `comment` FROM `order` WHERE `id` = '" & ID & "') AS `s`"
    sql = sql & " LEFT JOIN `table` AS `t` ON `s`.`tableid` = `t`.`id`"
    sql = sql & " LEFT JOIN `account` AS `a` ON `s`.`creatorid` = `a`.`id`;"
    
    Set MySQL = Nothing
    Set MySQL = New MYSQLDBLib.MySQL
    MySQL.CloseConnection
    CnMYSQL = MySQL.OpenConnection(Trim(vIPAddress), Trim(Str(vServerPort)), Trim(vLogin), Trim(vPassword), Trim(vDatabase), 0)
    If CnMYSQL < 0 Then
        frmDisplay.Pprint MsgerrCon & Trim(vIPAddress) & "!"
        GoTo err
    End If
    
    RS_Status = MySQL.OpenRecordset("tmp", sql)
    If (RS_Status = 0) Then
        frmDisplay.Pprint MsgerrQuery & "" & Trim(MySQL.LastError)
        GoTo err
    End If
    MySQL.MoveFirst "tmp"
    Do While (MySQL.EOR("tmp") = 0)
        If (PrintPreCheckZebra(ID, MySQL.GetFieldByName("tmp", "table"), MySQL.GetFieldByName("tmp", "creator"), MySQL.GetFieldByName("tmp", "comment"), MySQL.GetFieldByName("tmp", "date")) = True) Then
            PrintSucces ID
            frmDisplay.Pprint ("Check Printed " & ID)
        Else
            frmDisplay.Pprint ("Check Not Printed " & ID)
        End If
        MySQL.MoveNext "tmp"
    Loop
    MySQL.CloseConnection
    Exit Function
err:
MySQL.CloseConnection
frmDisplay.Pprint "Error: " & err.Description
End Function

Function PrintPreCheckZebra(order As String, table As String, creator As String, comment As String, ddate As String) As Boolean
Dim name_ua As String, price As Double, rprice As Double, col As Integer, count As Double, rcount As Double
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, TempCost As Double, TempRealCost As Double, TempAr(1 To 10) As String, TempPrice(1 To 10) As Double
Dim i As Integer, z As Integer, X As Integer, d_temp As String, SpecPrice As Boolean, HasDiscount As Boolean, tmp_sql As String
    SpecPrice = False 'если есть дисконт или предоплата
    HasDiscount = False 'если есть дисконт
    '---------------------------------------------------------------------------------------
    'Подготовка переменных для протокола
    '---------------------------------------------------------------------------------------
    Set MySQL = Nothing
    Set MySQL = New MYSQLDBLib.MySQL
    MySQL.CloseConnection
    CnMYSQL = MySQL.OpenConnection(Trim(vIPAddress), Trim(Str(vServerPort)), Trim(vLogin), Trim(vPassword), Trim(vDatabase), 0)
    If CnMYSQL < 0 Then
        PrintPreCheckZebra = False
        Exit Function
    End If
    '---------------------------------------------------------------------------------------
    'Печать шапки чека
    '--------------------
    Call BeginPrintText
    TempCost = 0
    PrintText String(29, "=")
    PrintText "Рахунок № " + order
    PrintText "Стiл № " + table
    PrintText "Офiцiант: " + creator
    PrintText "Дата: " + ddate
    'Разбиение и печать комментария (если он есть)
    '---------------------
    If comment <> "" Then
        If Len(comment) > 28 Then
            If (Len(comment) Mod 28) > 0 Then
                i = (Len(comment) \ 28) + 1
            Else
                i = (Len(comment) \ 28)
            End If
            X = 1
            For z = 1 To i
                If z = 1 Then
                    PrintText "Коментар: " + Mid(comment, X, 18)
                    X = X + 18
                    PrintText Mid(comment, X, 28)
                Else
                    If Mid(comment, X, 28) <> "" Then
                        PrintText Mid(comment, X, 28)
                    End If
                End If
                X = X + 28
            Next z
        Else
            PrintText "Коментар: " + comment
        End If
    End If
    
    'Получение рекордсета с пречеком
    tmp_sql = "SELECT `name_ua` as `name`, `cost` as `real_cost`, `discount_cost` as `cost`, `count` FROM account_ordered_item WHERE `orderid` = '" & order & "';"
    RS_Status = MySQL.OpenRecordset("tmp", tmp_sql)
    If (RS_Status = 0) Then
        frmDisplay.Pprint MsgerrQuery & "" & Trim(MySQL.LastError)
        PrintPreCheckZebra = False
        Exit Function
    End If
    
    MySQL.MoveFirst "tmp"
    Do While (MySQL.EOR("tmp") = 0)
        d_temp = InStr(CStr(MySQL.GetFieldByName("tmp", "name")), "ДИСКОН")
        If d_temp > 0 Then
            HasDiscount = True
            Temp1 = MySQL.GetFieldByName("tmp", "name")
            PrintText "Дисконт: " + Mid(Temp1, 18, 32)
            GoTo dok
        Else
            MySQL.MoveNext "tmp"
        End If
    Loop
dok:
    PrintText String(29, "=")
    'Конец шапки
    '===
    'Печать пунктов счета
    '---------------------
    i = 1
    MySQL.MoveFirst "tmp"
    Do While (MySQL.EOR("tmp") = 0)
            d_temp = InStr(CStr(MySQL.GetFieldByName("tmp", "name")), "ДИСКОН")
            If d_temp > 0 Then
                GoTo lo
            End If
            name_ua = Replace(CStr(MySQL.GetFieldByName("tmp", "name")), """", " ")
            price = Val(MySQL.GetFieldByName("tmp", "cost"))
            rprice = Val(MySQL.GetFieldByName("tmp", "real_cost"))
            col = Val(MySQL.GetFieldByName("tmp", "count"))
            count = col * price
            rcount = col * rprice
            TempCost = TempCost + count
            TempRealCost = TempRealCost + (col * rprice)
            'если внесена предоплата сохраняем данные
            If col < 0 Then
                TempAr(i) = name_ua + Space(29 - Len(name_ua + Format(count, "Standard"))) + Format(count, "Standard")
                TempPrice(i) = count
                TempCost = TempCost - count 'отнимаем это значение от суммы, т.к. она посчиталась выше
                i = i + 1
                GoTo lo
            End If
            'если товар не предоплата, печатаем
            Temp1 = name_ua
            If (rprice > price) Then
                Temp2 = Str(col) + " x " + Format(rprice, "Standard") + " = " + Format(rcount, "Standard")
                Temp3 = "Знижка: " + Format(100 - (price * 100 / rprice), "#0.00") + "%"
                Temp4 = Format(rprice - price, "Standard")
            Else
                Temp2 = Str(col) + " x " + Format(price, "Standard") + " = " + Format(count, "Standard")
            End If
            'Если название не влезает - режем на 2 чати
            If Len(Temp1) > 29 Then
                PrintText Mid(Temp1, 1, 29)
                PrintText Mid(Temp1, 30, 58)
            Else
                PrintText Temp1
            End If
            If (rprice > price) Then
                PrintText Space(29 - Len(Temp2)) + Temp2
                PrintText Temp3 + Space(29 - (Len(Temp3) + Len(Temp4))) + Temp4
                PrintText Space(29)
            Else
                PrintText Space(29 - Len(Temp2)) + Temp2
                PrintText Space(29)
            End If
lo:     MySQL.MoveNext "tmp"
    Loop
    PrintText String(29, "=")
    Temp1 = "~СУМА: "
    Temp2 = Format(TempRealCost, "Standard")
    PrintText Temp1 + Space(26 - Len(Temp1 + Temp2)) + Temp2
    '=====================================
    'если есть дисконтная карта
    If HasDiscount Or TempCost <> TempRealCost Then
        Temp1 = "ЗНИЖКА: "
        Temp2 = Format((TempCost - TempRealCost) * -1, "Standard")
        PrintText Temp1 + Space(29 - Len(Temp1 + Temp2)) + Temp2
        SpecPrice = True
    End If
    '=====================================
    'если есть данные по предоплате
    If TempAr(1) <> "" Then
        PrintText String(29, "-")
        For z = 1 To i
            If TempAr(z) <> "" Then
                PrintText TempAr(z)
                TempCost = TempCost + TempPrice(z)
            End If
        Next z
        SpecPrice = True
    End If
    If SpecPrice = True Then
        Temp1 = "~ДО СПЛАТЫ: "
        Temp2 = Format(TempCost, "Standard")
        PrintText Temp1 + Space(26 - Len(Temp1 + Temp2)) + Temp2
    End If
    PrintText String(29, "=")
    PrintText Now()
    PrintText " "
    '=====================================
    MySQL.CloseConnection
    If (PrintPreCheck(LoadPrinter(), order) = False) Then
        PrintPreCheckZebra = False
    Else
        PrintPreCheckZebra = True
    End If
End Function


Function PrintSucces(ID As String) As Boolean
    On Error GoTo err
    Set MySQL = Nothing
    Set MySQL = New MYSQLDBLib.MySQL
    MySQL.CloseConnection
    CnMYSQL = MySQL.OpenConnection(Trim(vIPAddress), Trim(Str(vServerPort)), Trim(vLogin), Trim(vPassword), Trim(vDatabase), 0)
    If CnMYSQL < 0 Then
        GoTo err
    End If
    PrintSucces = False
    '------------------------------------------------------------------
    temp_sql = "UPDATE `account_ordered_item` SET `isPrinted` = 1 WHERE `orderid` = '" & ID & "'"
    RS_Status = MySQL.Execute(temp_sql)
    If (RS_Status = 0) Then
        frmDisplay.Pprint MsgerrQuery & "" & Trim(MySQL.LastError)
        GoTo err
    End If
    '------------------------------------------------------------------
    PrintSucces = True
    MySQL.CloseConnection
    Exit Function
err:
    MySQL.CloseConnection
    PrintSucces = False
End Function
