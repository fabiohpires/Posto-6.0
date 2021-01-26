Module EditaINI
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    '****************************************************
    'Windows API/Global Declarations for :Multiline .INI
    '****************************************************

    Public Function ReadINI(ByRef Section As String, ByRef KeyName As String, ByRef Default_Renamed As String, ByRef filename As String) As String
        Dim sRet As String
        sRet = New String(Chr(0), 1000)
        ReadINI = Left(sRet, GetPrivateProfileString(Section, KeyName, Default_Renamed, sRet, Len(sRet), filename))
        ReadINI = Left(sRet, InStr(sRet, Chr(0)))
    End Function

    Public Function WriteINI(ByRef sSection As String, ByRef sKeyName As String, ByRef sNewString As String, ByRef sFileName As String) As Short
        Dim r As Integer
        r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
    End Function

End Module
