Attribute VB_Name = "INI"
Option Explicit



Public Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
' This first line is the declaration from win32api.txt

Public Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Public Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Public Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Function VBGetPrivateProfileString(section$, key$, File$) As String
    Dim KeyValue$
    Dim characters As Long
    
    
    KeyValue$ = String$(128, 0)
    characters = GetPrivateProfileStringByKeyName(section$, key$, "", KeyValue$, 127, File$)

    If characters > 1 Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    
    VBGetPrivateProfileString = KeyValue$

End Function


Public Function GetSectionNames(FileName As String, SectionNames As Variant) As Integer
    'GetSectionNames Return Number of Section in file
    'SectionNames return all section names
    
    Dim characters As Long
    Dim SectionList As String
    Dim ArrSection() As String
    Dim i As Integer
    Dim NullOffset%
    
    SectionList = String$(128, 0)

    ' Retrieve the list of keys in the section
    characters = GetPrivateProfileStringSections(0, 0, "", SectionList, 127, FileName)
    
    ' Load sections into Arrey
    i = 0
    Do
        NullOffset% = InStr(SectionList, Chr$(0))
        If NullOffset% > 1 Then
            ReDim Preserve ArrSection(i)
            ArrSection(i) = Mid$(SectionList, 1, NullOffset% - 1)
            SectionList$ = Mid$(SectionList, NullOffset% + 1)
            i = i + 1
        End If
    Loop While NullOffset% > 1
    GetSectionNames = i - 1
    SectionNames = ArrSection
    

End Function

Public Function GetKeyNames(SectionName As String, FileName As String, KeyNames As Variant) As Integer
    'GetKeyNames Return Number of key in section
    'KeyNames Return list of keyNames in section
    
    Dim characters As Long
    Dim KeyList As String
    Dim ArrKey() As String
    Dim i As Integer
    
    KeyList = String$(128, 0)
    ' Retrieve the list of keys in the section
    
    characters = GetPrivateProfileStringKeys(SectionName, 0, "", KeyList, 127, FileName)
    
    ' Load Keys into Arrey
    Dim NullOffset%
    i = 0
    Do
        NullOffset% = InStr(KeyList, Chr$(0))
        If NullOffset% > 1 Then
            ReDim Preserve ArrKey(i)
            ArrKey(i) = Mid$(KeyList, 1, NullOffset% - 1)
            KeyList$ = Mid$(KeyList, NullOffset% + 1)
            i = i + 1
        End If
    Loop While NullOffset% > 1
    GetKeyNames = i - 1
    KeyNames = ArrKey
End Function

Public Function DeleteKey(KeyName As String, SectionName As String, FileName As String) As Long
    'Return 0 if Deletion not sucsesful
       
    ' Delete the selected key
    'DeleteKey = WritePrivateProfileStringToDeleteKey(SecName, lstKeys.Text, 0, FileName$)

    
End Function

Public Function WriteKey(SectionName As String, KeyName As String, KeyValue As String, FileName As String) As Long
    If Len(KeyValue) = 0 Then KeyValue = " "
    WriteKey = WritePrivateProfileStringByKeyName(SectionName, KeyName, KeyValue, FileName)
End Function

Public Function WriteSection(SectionName As String, FileName As String) As Long
    WriteSection = WritePrivateProfileSection(SectionName, "", FileName)
End Function
    
    
Public Function DeleteSection(SectionName, FileName) As Long
    DeleteSection = WritePrivateProfileStringToDeleteSection(SectionName, 0&, 0&, FileName)
End Function
    
