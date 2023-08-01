Attribute VB_Name = "WRini"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal w_returnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'----------------------------------------------------
'------------從INI文件中取得數據---------------------
'---設計者:,設計日期:2004/04/07----------------
'----------------------------------------------------
Public Function GetIni(ByVal KeyName As String, Optional ByVal KeyGroup As String, Optional ByVal PathName As String) As Variant
    Dim strKeyGroup As String, strPathName As String, strRtn As String, nSize As Long
    Dim I As Long
    strRtn = String$(255, Chr(0))
    strKeyGroup = IIf(KeyGroup = "", "Setup", KeyGroup)
    If PathName = "" Then
        strPathName = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "set.ini"
    Else
        strPathName = PathName
    End If
    nSize = GetPrivateProfileString(strKeyGroup, KeyName, "", strRtn, Len(strRtn), strPathName)

    If nSize > 0 Then
        GetIni = Left$(strRtn, nSize)
        I = InStr(GetIni, Chr(0))
        If I > 0 Then GetIni = Left(GetIni, I - 1)
    End If
End Function

Public Function uLen(ByVal strBuf As String) As Long
    uLen = LenB(StrConv(strBuf, vbFromUnicode))
End Function

'----------------------------------------------------
'------------寫入數據到INI文件中---------------------
'---設計者:,設計日期:2004/04/07----------------
'----------------------------------------------------
Public Function SetIni(ByVal KeyValue As String, ByVal KeyName As String, Optional ByVal KeyGroup As String, Optional ByVal PathName As String) As Long
    Dim strKeyGroup As String, strPathName As String
    strKeyGroup = IIf(KeyGroup = "", "Setup", KeyGroup)
    If PathName = "" Then
        strPathName = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Mcs.ini"
    Else
        strPathName = PathName
    End If
    
    SetIni = WritePrivateProfileString(strKeyGroup, KeyName, KeyValue, strPathName)
End Function
Sub SetiniUser(Name As String, Value As String, File As String)
    SetIni Value, Name, "账户", App.Path & "\data\" & File & "\" & File & ".ini"
End Sub
Function ReadIniUser(Name As String, File As String) As String
    ReadIniUser = GetIni(Name, "账户", App.Path & "\data\" & File & "\" & File & ".ini")
End Function
