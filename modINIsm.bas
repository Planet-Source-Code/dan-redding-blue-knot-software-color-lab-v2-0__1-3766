Attribute VB_Name = "modINI"
Option Explicit
Declare Function GetPrivateProfileStringByKey _
    Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileStringKeys _
    Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Long, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileStringSections _
    Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileInt _
    Lib "kernel32" Alias "GetPrivateProfileIntA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal nDefault As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileStringByKey _
    Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpString As String, ByVal lpFileName As String) As Long
    
Declare Function WritePrivateProfileStringToDeleteKey _
    Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
    ByVal lpString As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection _
    Lib "kernel32" Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileSection _
    Lib "kernel32" Alias "WritePrivateProfileSectionA" _
    (ByVal lpAppName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Public Sub DeleteINIKey(Section As String, Key As String)
Dim lReturn As Long
'deletes entire entry, not just value
    lReturn = WritePrivateProfileStringToDeleteKey(Section, Key, 0&, "drqcr.ini")
    
End Sub

Public Sub WriteINI(Section As String, Key As String, ByVal KeyValue As String)
Dim lReturn As Long
'writes value (creates entry if it doesn't exist)
    lReturn = WritePrivateProfileStringByKey(Section, Key, KeyValue, "drqcr.ini")
End Sub

Public Function GetINIString(Section As String, Key As String, Optional Default As String = "") As String
Dim KeyValue As String, Characters As Long, intPos As Integer
'retrieves STRING value
    KeyValue = String(256, 0)
    
    Characters = GetPrivateProfileStringByKey(Section, Key, Default, KeyValue, 255, "drqcr.ini")
    
    If Characters > 1 Then KeyValue = left$(KeyValue, Characters)
    
    intPos = InStr(KeyValue, Chr$(0) & Chr$(0))
    
    If intPos > 0 Then
        KeyValue = left$(KeyValue, intPos - 1)
    End If
    
    GetINIString = KeyValue

End Function

Public Function GetININumber(Section As String, Key As String, Optional Default As Long = 0) As Long
'retrieves numeric value
    GetININumber = GetPrivateProfileInt(Section, Key, Default, "drqcr.ini")
End Function
