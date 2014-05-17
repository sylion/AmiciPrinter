Attribute VB_Name = "ZebraAPI"
Option Explicit
Option Compare Text

Public Const MAX_LINELEN& = 29
Private Const Start_Y# = 110
Private Const Step_Y# = 30
Private Const Step_G_Y# = 50
Private Const End_Y# = 2100

'-----------------------------------------------------------------------
Public Check_Array() As String
'-----------------------------------------------------------------------

Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_LOCAL = &H2

Public Type PRINTER_INFO_1
    FLAGS As Long
    pDescription As String
    PName As String
    PComment As String
End Type

Public Type PRINTER_INFO_4
    pPrinterName As String
    pServerName As String
    Attributes As Long
End Type

Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

'--------------------------------------------------------------------------
Public Declare Function EnumPrinters Lib "winspool.drv" Alias _
         "EnumPrintersA" (ByVal FLAGS As Long, ByVal Name As String, _
         ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, _
         pcbNeeded As Long, pcReturned As Long) As Long
      
'--------------------------------------------------------------------------
Public Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
         (ByVal RetVal As String, ByVal Ptr As Long) As Long

'--------------------------------------------------------------------------
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenA" _
         (ByVal Ptr As Long) As Long

'--------------------------------------------------------------------------
Public Declare Sub Sleep Lib "kernel32" _
         (ByVal dwMilliseconds As Long)

'--------------------------------------------------------------------------
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
        hPrinter As Long) As Long
  
'--------------------------------------------------------------------------
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
        hPrinter As Long) As Long
  
'--------------------------------------------------------------------------
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
        hPrinter As Long) As Long
  
'--------------------------------------------------------------------------
Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
        "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
        ByVal pDefault As Long) As Long
  
'--------------------------------------------------------------------------
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias _
        "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
        pDocInfo As DOCINFO) As Long
  
'--------------------------------------------------------------------------
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
        hPrinter As Long) As Long
  
'--------------------------------------------------------------------------
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal _
        hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
        pcWritten As Long) As Long

'----------------------------------------------------------------
' ?????? ????????
'----------------------------------------------------------------
Function PrintPreCheck(NamePrinter As String, OrderId As String) As Boolean
    Dim Set_Zebra As String
    Dim End_Zebra As String
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim MyDocInfo As DOCINFO
    Dim sWrittenData As String
    Dim SplitData() As String
    Dim Item As String
    Dim CountItem As String
    Dim lDoc As Long
    Dim lpcWritten As Long
    Dim i As Long
    Dim Out_Str As String
    Dim YY As Long
    Dim flag_end As Boolean
    On Error GoTo err
    PrintPreCheck = False
    
    Dim Logo_Print As String
    Logo_Print = "GG0,0,""LOGO"""
    '--------------------------------------------------------------------------
    Set_Zebra = "N" & vbCrLf
    End_Zebra = "P1"
    '--------------------------------------------------------------------------
    lReturn = OpenPrinter(NamePrinter, lhPrinter, 0)
    If lReturn = 0 Then
        frmDisplay.Pprint "Printer " & NamePrinter & " Not found!"
        GoTo err
    End If
    '--------------------------------------------------------------------------
    MyDocInfo.pDocName = "Заказ № " & CStr(OrderId)
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    '-----------------------------------------------------------------------
    sWrittenData = Set_Zebra & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                   Len(sWrittenData), lpcWritten)
    '-----------------------------------------------------------------------
    sWrittenData = Logo_Print & vbCrLf
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                   Len(sWrittenData), lpcWritten)
    '-----------------------------------------------------------------------
    YY = Start_Y
    flag_end = False
    If UBound(Check_Array) > 0 Then
        For i = 1 To UBound(Check_Array)
            If Left(Check_Array(i), 1) = "~" Then
                Out_Str = "A0," & Trim(CStr(YY)) & ",0,4,1,2,N,"""
                Out_Str = Out_Str & Replace(Check_Array(i), "~", "") & """"
                YY = YY + Step_G_Y
            Else
                If Left(Check_Array(i), 1) = "^" Then
                    Check_Array(i) = Replace(Check_Array(i), "^", "")
                    '-----------------------------------------------------------------------
                    SplitData = Split(Check_Array(i), Chr(9))
                    Item = Trim(Left(SplitData(0), 50))
                    CountItem = Trim(SplitData(1))
                    '-----------------------------------------------------------------------
                    Out_Str = "A0," & Trim(CStr(YY + 5)) & ",0,4,1,2,N,"""
                    Out_Str = Out_Str & Space(25 - Len(CountItem)) & CountItem & """"
                    Out_Str = Out_Str & vbCrLf
                    '-----------------------------------------------------------------------
                    If (Len(Item) < 26) Then
                        Out_Str = Out_Str & "A0," & Trim(CStr(YY + Step_Y / 2)) & ",0,3,1,1,N,"""
                        Out_Str = Out_Str & Item & """"
                        YY = YY + Step_Y * 2
                    Else
                        Out_Str = Out_Str & "A0," & Trim(CStr(YY)) & ",0,3,1,1,N,"""
                        Out_Str = Out_Str & Trim(Left(Item, 25)) & """"
                        YY = YY + Step_Y
                        Out_Str = Out_Str & vbCrLf
                        Out_Str = Out_Str & "A0," & Trim(CStr(YY)) & ",0,3,1,1,N,"""
                        Out_Str = Out_Str & Trim(Mid(Item, 26, 25)) & """"
                        YY = YY + Step_Y
                    End If
                    '-----------------------------------------------------------------------
                Else
                    Out_Str = "A0," & Trim(CStr(YY)) & ",0,3,1,1,N,"""
                    Out_Str = Out_Str & Check_Array(i) & """"
                    YY = YY + Step_Y
                End If
            End If
            '-----------------------------------------------------------------------
            sWrittenData = Out_Str & vbCrLf
            lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                      Len(sWrittenData), lpcWritten)
            '-----------------------------------------------------------------------
            frmDisplay.Pprint Out_Str
            '-----------------------------------------------------------------------
            If YY > End_Y Then
                YY = Start_Y
                '-----------------------------------------------------------------------
                sWrittenData = End_Zebra & vbCrLf & "N" & vbCrLf
                lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                          Len(sWrittenData), lpcWritten)
                flag_end = True
            Else
                flag_end = False
            End If
        Next i
        If flag_end = False Then
            '-----------------------------------------------------------------------
            sWrittenData = End_Zebra & vbCrLf
            lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                      Len(sWrittenData), lpcWritten)
            '-----------------------------------------------------------------------
        End If
    End If
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    '-----------------------------------------------------------------------
    PrintPreCheck = True
    Exit Function
err:
    PrintPreCheck = False
End Function

Sub BeginPrintText()
    ReDim Check_Array(0)
End Sub

Sub PrintText(comment As String)
    Dim i As Long
    i = UBound(Check_Array) + 1
    ReDim Preserve Check_Array(i)
    Check_Array(i) = Replace(comment, """", "'")
End Sub


