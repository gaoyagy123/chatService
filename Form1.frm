VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "聊天服务器端"
   ClientHeight    =   3810
   ClientLeft      =   8820
   ClientTop       =   5100
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2280
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2640
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function getsockopt Lib "WS2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByRef optlen As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'定义常量
Const BUSY As Boolean = False
Const FREE As Boolean = True
'定义连接状态
Private Const MAX_LVMSTRING As Long = 512
Private Const MEM_COMMIT = &H1000&
Private Const MEM_RELEASE = &H8000&
Private Const PAGE_READWRITE = &H4&

Private Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type PicRev
    Start As Boolean
    Target As String
    MaxLen As Long
    RevLen As Long
    ByteAddr As Long
End Type
Private Type PicInfo
    MaxLen As Long
    Sel As Long
End Type
Private Type SendHead
    head  As String * 4
    UserMsg As String * 20
    UserLen As Long
    MaxLenth As Long
    ObjNum As Long
    Pic(99) As PicInfo
End Type
Private Type msgHead
    head  As String * 4
    UserMsg As String * 20
    UserLen As Long
    MaxLenth As Long
End Type

Dim PicSend(999) As PicRev
Dim ConnectState() As Boolean, Connected() As Date
Dim Conn As New ADODB.Connection


Private Sub Command1_Click()
    Dim Rs As New ADODB.Recordset, I As Long, Sxb() As Byte
    Rs.Open "Select * from userdata", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    Rs.MoveFirst
    Do While Not Rs.EOF
        Sxb = GetSendByte("SyS", "System", Text2.text)
        I = GetTargetIndex(Rs(1))
        If I >= 0 Then
            Winsock1(I).SendData Sxb
        Else
            SetSystemData CStr(Rs(0)), Sxb
        End If
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub Form_Load()
    ReDim Preserve ConnectState(0 To 0)
    ReDim Preserve Connected(0 To 0)
    On Error Resume Next
    ConnectState(0) = FREE
    Connected(0) = Now
    '指定网络端口号
    listener.LocalPort = "5556"
    '开始侦听
    listener.Listen
    Timer1.Enabled = True
    InitData
End Sub
Sub InitData()
    Dim CnnStr  As String
    CnnStr = "DRIVER={MySQL ODBC 5.3 ANSI Driver};server=localhost;port=3307;uid=root;pwd=123456;database=mytalk"
    Conn.Open CnnStr
    Conn.CursorLocation = adUseClient
End Sub


Private Sub Form_Unload(Cancel As Integer)
    listener.Close
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
    Dim SockIndex As Integer
    Dim SockNum As Integer
    On Error Resume Next
    Dtxt requestID & "连接请求"
    '查找连接的用户数
    SockNum = UBound(ConnectState)
    If SockNum > 14 Then
       ' Exit Sub
    End If
    
    '查找空闲的sock
    SockIndex = FindFreeSocket
    
    '如果已有的sock都忙，而且sock数不超过15个，动态添加sock
    If SockIndex > SockNum Then
        Load Winsock1(SockIndex)
    End If
    ConnectState(SockIndex) = BUSY
    Connected(SockIndex) = Now
    
    '接受请求
    Winsock1(SockIndex).Accept requestID
End Sub


Private Sub Timer1_Timer()
    Dim I As Integer, Dxb() As Byte
    For I = 0 To UBound(ConnectState)
        If ConnectState(I) = BUSY Then
            If DateDiff("s", Connected(I), Now) > 10 Or Winsock1(I).State = 9 Then
                Winsock1_Close I
            Else
                Debug.Print "状态：" & Winsock1(I).State
                Dxb = GetSendByte("Chk", "System", "")
                If Winsock1(I).State = sckConnected Then Winsock1(I).SendData Dxb
            End If
        End If
    Next
End Sub

'客户断开，关闭相应的sock
Private Sub Winsock1_Close(Index As Integer)
    Dtxt Winsock1(Index).LocalIP & "断开了"
    If Winsock1(Index).State <> sckClosed Then
        Winsock1(Index).Close
    End If
    ConnectState(Index) = FREE
    Winsock1(Index).Tag = ""
End Sub

'接收数据
Sub GetCache(mvarSocketHandle As Long)
    Dim DataSize As Long
    rd = getsockopt(mvarSocketHandle, &HFFFF&, &H1002&, DataSize, 4)

   Debug.Print DataSize, 333333333
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dtxt "数据来自" & Winsock1(Index).LocalIP

    Dim Dxb() As Byte, Rec As Long, Txb() As Byte

    Winsock1(Index).GetData Dxb, vbByte
    
    Txb = Dxb
    Debug.Print UBound(Dxb) + 1
    Do
        If Not PicSend(Index).Start Then
            Rec = DoNomral(Index, Dxb)
        Else
            Rec = PicSendUser(Index, Dxb)
        End If
        
        If Rec > 0 Then
            ReDim Dxb(Rec - 1)
            CopyMemory Dxb(0), Txb(UBound(Txb) + 1 - Rec), Rec
        End If
    Loop While Rec > 0
End Sub

Sub PicSendUser2(Index As Integer, Dxb() As Byte)
    Winsock1(PicSend(Index).Target).SendData Dxb
    PicSend(Index).RevLen = PicSend(Index).RevLen + UBound(Dxb) + 1
    Debug.Print PicSend(Index).RevLen, PicSend(Index).MaxLen, 2222
    If PicSend(Index).RevLen >= PicSend(Index).MaxLen Then
        PicSend(Index).Start = False
    End If
End Sub

Sub LoginUser(ByVal Index As Integer, Dxb() As Byte)
    Dim Uname As String, Passwd As String, tmp As String
    Dim pHead As msgHead
    Dim Info As String, numID As Long, Sxb() As Byte
    Dim I As Long
    
    CopyMemory pHead, Dxb(0), Len(pHead)
    Uname = pHead.UserMsg
    Uname = Left(Uname, pHead.UserLen)
    Passwd = String$(pHead.MaxLenth, Chr(0))
    CopyMemory ByVal Passwd, Dxb(Len(pHead)), Len(Passwd)
    If CheckData(Uname, Passwd) Then
        I = GetTargetIndex(Uname)
        If I >= 0 Then
            Sxb = GetSendByte("LoF", "System", "")
            Winsock1(Index).SendData Sxb
            Exit Sub
        End If
        Winsock1(Index).Tag = Uname
        Info = LoginMsg(Uname)
        Sxb = GetSendByte("Int", "System", Info)
        Winsock1(Index).SendData Sxb
        GetData Uname, numID
        GetSystemData numID, Index
        'GetUserData numID, Index
    Else
        Winsock1_Close Index
    End If
End Sub
Function CheckData(Name As String, Passwd As String) As Boolean
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from userdata where ID=""" & Name & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    If Rs("password") = Passwd Then CheckData = True
    Rs.Close
    Set Rs = Nothing
End Function
Function GetSendByte(head As String, User As String, Info As String) As Byte()
    Dim mhead As msgHead
    Dim Dxb() As Byte, size As Long, hLen As Long
    hLen = Len(mhead)
    mhead.head = head
    mhead.UserMsg = User
    mhead.UserLen = uLen(User)
    mhead.MaxLenth = uLen(Info)
    
    size = hLen + uLen(Info)
    ReDim Dxb(size - 1)
    CopyMemory Dxb(0), mhead, hLen
    If uLen(Info) > 0 Then CopyMemory Dxb(hLen), ByVal Info, uLen(Info)
    
    GetSendByte = Dxb
End Function
Function DoNomral(Index As Integer, Dxb() As Byte) As Long
    Dim pHead As msgHead
    CopyMemory pHead, Dxb(0), Len(pHead)
    DoNomral = UBound(Dxb) - Len(pHead) - pHead.MaxLenth + 1
    If Left(pHead.head, 3) = "Lgi" Then
        LoginUser Index, Dxb
    ElseIf Left(pHead.head, 3) = "Tak" Then
        UserToUser Index, Dxb
    ElseIf Left(pHead.head, 3) = "Ico" Then
        Dim Dx As String, size As Long, L As Long
        L = Len(pHead)
        size = UBound(Dxb) + 1 - L
        Dx = String$(size, Chr(0))
        CopyMemory ByVal Dx, Dxb(L), size
        SetData Winsock1(Index).Tag, "Icon", Val(Dx)
    ElseIf Left(pHead.head, 3) = "Rek" Then
        Connected(Index) = Now
    ElseIf Left(pHead.head, 3) = "Pic" Then
        GetPicTarget Index, Dxb
        Dim head As SendHead
        DoNomral = UBound(Dxb) - Len(head) - pHead.MaxLenth + 1
    ElseIf Left(pHead.head, 3) = "Fid" Then
        GetFind Index, Dxb
    ElseIf Left(pHead.head, 3) = "AdF" Then
        AddFriend Index, Dxb
    ElseIf Left(pHead.head, 3) = "AdR" Then
        RecvieFriend Index, Dxb
    ElseIf Left(pHead.head, 3) = "DeF" Then
        DelFriend Index, Dxb
    End If
End Function
Sub RecvieFriend(Index As Integer, Dxb() As Byte)
    Dim Dx As String, size As Long, L As Long, tmp() As String
    Dim Res As String, Name As String, myName As String
    Dim pHead As msgHead
    Dim I As Long, Icon As Long, ID As Long, Nick As String, Fr As String
    
    L = Len(pHead)
    size = UBound(Dxb) + 1 - L
    Dx = String$(size, Chr(0))
    CopyMemory ByVal Dx, Dxb(L), size
    CopyMemory pHead, Dxb(0), Len(pHead)
    
    myName = pHead.UserMsg
    myName = Left(myName, pHead.UserLen)
        
    tmp = Split(Dx, "|")
    If tmp(1) = "1" Then
        GetDataFromID tmp(0), Name, , Nick, , , , Icon
        AddFriendData Name, myName
        If Name <> myName Then
            GetData myName, ID, , Nick, , , , Icon
            AddFriendData myName, Name
        End If
    End If
    pHead.head = "AdR"
    I = GetTargetIndex(Name)
    If I <> Index Then
        Res = GetUserMsg("ID", myName)
        Res = Res & "|" & tmp(1)
        pHead.MaxLenth = uLen(Res)
        
        ReDim Dxb(L + uLen(Res) - 1)
        CopyMemory Dxb(0), pHead, L
        CopyMemory Dxb(L), ByVal Res, uLen(Res)
        
        If I >= 0 Then
            Winsock1(I).SendData Dxb
        Else
            SetSystemData tmp(0), Dxb
        End If
    End If
    If tmp(1) = "1" Then
        pHead.UserMsg = Name
        pHead.UserLen = uLen(Name)
        Res = GetUserMsg("ID", Name)
        Res = Res & "|" & tmp(1)
        pHead.MaxLenth = uLen(Res)
        
        ReDim Dxb(L + uLen(Res) - 1)
        CopyMemory Dxb(0), pHead, L
        CopyMemory Dxb(L), ByVal Res, uLen(Res)
        Winsock1(Index).SendData Dxb
    End If
End Sub
Sub AddFriend(Index As Integer, Dxb() As Byte)
    Dim Dx As String, size As Long, L As Long, tmp() As String
    Dim Res As String, Name As String, myName As String
    Dim pHead As msgHead
    Dim I As Long, Icon As Long, ID As Long, Nick As String
    
    L = Len(pHead)
    size = UBound(Dxb) + 1 - L
    Dx = String$(size, Chr(0))
    CopyMemory ByVal Dx, Dxb(L), size
    CopyMemory pHead, Dxb(0), Len(pHead)
    
    tmp = Split(Dx, "|")
    GetDataFromID tmp(0), Name
    
    I = GetTargetIndex(Name)
    myName = pHead.UserMsg
    myName = Left(myName, pHead.UserLen)
    Res = GetUserMsg("ID", myName)
    Res = Res & "|" & tmp(1)
    
    pHead.head = "AdF"
    pHead.MaxLenth = uLen(Res)
    
    ReDim Dxb(L + uLen(Res) - 1)
    CopyMemory Dxb(0), pHead, L
    CopyMemory Dxb(L), ByVal Res, uLen(Res)
    If I >= 0 Then
        Winsock1(I).SendData Dxb
    Else
        SetSystemData tmp(0), Dxb
    End If
End Sub
Sub GetFind(Index As Integer, Dxb() As Byte)
    Dim Dx As String, size As Long, L As Long
    Dim Res As String, Res1 As String, Res2 As String
    Dim p1() As String, p2() As String
    Dim pHead As msgHead
    
    L = Len(pHead)
    size = UBound(Dxb) + 1 - L
    Dx = String$(size, Chr(0))
    CopyMemory ByVal Dx, Dxb(L), size
    
    Res1 = GetUser("num", Dx, Winsock1(Index).Tag)
    p1 = Split(Res1, "|")
    Res2 = GetUser("nick", Dx, Winsock1(Index).Tag)
    p2 = Split(Res2, "|")
    
    Res = CheckDelData(p1, p2, Res1 & Res2)
    If Res = "" Then Res = "Faild"
    If Right(Res, 1) = "|" Then Res = Left(Res, Len(Res) - 1)
    
    pHead.head = "Fid"
    pHead.UserMsg = Winsock1(Index).Tag
    pHead.UserLen = uLen(Winsock1(Index).Tag)
    pHead.MaxLenth = uLen(Res)
    
    ReDim Dxb(L + uLen(Res) - 1)
    CopyMemory Dxb(0), pHead, L
    CopyMemory Dxb(L), ByVal Res, uLen(Res)
    Winsock1(Index).SendData Dxb
End Sub
Sub DelFriend(Index As Integer, Dxb() As Byte)
    Dim Dx As String, L As Long
    Dim pHead As msgHead, Name As String
    
    L = Len(pHead)

    CopyMemory pHead, Dxb(0), Len(pHead)
    
    Dx = String$(pHead.MaxLenth, Chr(0))
    Name = pHead.UserMsg
    Name = Left(Name, pHead.UserLen)
    CopyMemory ByVal Dx, Dxb(L), pHead.MaxLenth
    
    Dim Rs As New ADODB.Recordset
    Rs.Open "delete from friend where idfriend1=""" & Name & """ and idfriend2=""" & Dx & """", Conn, adOpenDynamic, 3
    Rs.Open "delete from friend where idfriend1=""" & Dx & """ and idfriend2=""" & Name & """", Conn, adOpenDynamic, 3
    Set Rs = Nothing
    
    Dim Sxb() As Byte, Info As String, I As Long
    Info = LoginMsg(Name)
    Sxb = GetSendByte("Int", "System", Info)
    Winsock1(Index).SendData Sxb

    I = GetTargetIndex(Dx)
    If I >= 0 Then
        Info = LoginMsg(Dx)
        Sxb = GetSendByte("Int", "System", Info)
        Winsock1(I).SendData Sxb
    End If
End Sub

Function CheckDelData(p1() As String, p2() As String, Res As String) As String
    Dim I As Long, J As Long
    If SafeArrayGetDim(p1) = 0 Or SafeArrayGetDim(p2) = 0 Then
        CheckDelData = Res
        Exit Function
    End If
    For I = 0 To UBound(p1)
        For J = 0 To UBound(p2)
            If p1(I) = p2(J) And Len(p1(I)) > 0 Then
                Res = Replace(Res, p1(I) & "|", "", , 1)
            End If
        Next
    Next
    CheckDelData = Res
End Function
Function GetUser(ByVal Group As String, ByVal Value As String, Optional myName As String) As String
    Dim Rs As New ADODB.Recordset
    If Group = "num" Then Value = Val(Value)
    Rs.Open "Select * from userdata where " & Group & " like ""%" & Value & "%""", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    Rs.MoveFirst
    Do While Not Rs.EOF
        If Len(myName) = 0 Or Not CheckFriendExists(myName, Rs("ID")) Then
            GetUser = GetUser & Rs("ID") & "," & Rs("nick") & "," & Rs("num") & "," & Rs("Icon") & "|"
        End If
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End Function
Function GetUserMsg(Group As String, Value As String) As String
    Dim Rs As New ADODB.Recordset
    If Group = "num" Then Value = Val(Value)
    Rs.Open "Select * from userdata where " & Group & " = """ & Value & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    GetUserMsg = Rs("ID") & "," & Rs("nick") & "," & Rs("num") & "," & Rs("Icon")
    Rs.Close
    Set Rs = Nothing
End Function
Sub SetData(Name As String, Group As String, Value)
    Dim Rs As New ADODB.Recordset
    'Rs.Open "Select * from userdata where ID=""" & Name & """", Conn, adOpenDynamic, 3
   ' If Not Rs.EOF Then
   '     Rs(Group) = Value
  '      Rs.Update
  '  End If
  '  Rs.Close
    Rs.Open "Update userdata set Icon = """ & Value & """ where ID=""" & Name & """", Conn, adOpenDynamic, 3
    Set Rs = Nothing
End Sub
Sub AddFriendData(Name As String, Friends As String)

    If CheckFriendExists(Name, Friends) Then Exit Sub
    
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from friend", Conn, adOpenDynamic, 3
    Rs.AddNew
    Rs(0) = Name
    Rs(1) = Friends
    Rs.Update
    Rs.Close
    Set Rs = Nothing
End Sub
Function GetFriend(myName As String) As String()
    Dim Rs As New ADODB.Recordset, Fr() As String, I As Long
    Rs.Open "Select * from friend where idfriend1=""" & myName & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    ReDim Fr(Rs.RecordCount - 1) As String
    Rs.MoveFirst
    Do While Not Rs.EOF
        Fr(I) = Rs(1)
        I = I + 1
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
    GetFriend = Fr
End Function

Function CheckFriendExists(Name As String, Friends As String) As Boolean
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from friend where idfriend1=""" & Name & """ and idfriend2=""" & Friends & """", Conn, adOpenDynamic, 3
    CheckFriendExists = Not Rs.EOF
    Rs.Close
    Set Rs = Nothing
End Function

Sub SetSystemData(Name As String, Value() As Byte)
    Dim Rs As New ADODB.Recordset, I As Long
    Rs.Open "Select * from systempost", Conn, adOpenDynamic, 3
    Rs.AddNew
    Rs(0) = Name
    Rs(1) = Value
    Rs.Update
    Rs.Close
    Set Rs = Nothing
End Sub
Sub SetUserData(Name As String, Value() As Byte)
    Dim Rs As New ADODB.Recordset, I As Long
    Rs.Open "Select * from userpost", Conn, adOpenDynamic, 3
    Rs.AddNew
    Rs(0) = Name
    Rs(1) = Value
    Rs.Update
    Rs.Close
    Set Rs = Nothing
End Sub
Sub GetSystemData(numID As Long, Index As Integer)
    Dim Rs As New ADODB.Recordset, I As Long
    Dim Dxb() As Byte
    Rs.Open "Select * from systempost where usernum=""" & numID & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    Rs.MoveFirst
    Do While Not Rs.EOF
        Dxb = Rs(1)
        Winsock1(Index).SendData Dxb
        Rs.MoveNext
    Loop
    Rs.Close
    Rs.Open "delete from systempost where usernum=""" & numID & """", Conn, adOpenDynamic, 3
    Set Rs = Nothing
End Sub
Sub GetUserData(numID As Long, Index As Integer)
    Dim Rs As New ADODB.Recordset, I As Long
    Dim Dxb() As Byte
    Rs.Open "Select * from userpost where numid=""" & numID & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    Rs.MoveFirst
    Do While Not Rs.EOF
        Dxb = Rs(1)
        Winsock1(Index).SendData Dxb
        Rs.MoveNext
    Loop
    Rs.Close
    Rs.Open "delete from userpost where numid=""" & numID & """", Conn, adOpenDynamic, 3
    Set Rs = Nothing
End Sub
Function GetPicTarget2(Index As Integer, Dxb() As Byte) As String
    Dim User As String
    Dim Target As Integer
    Dim head As SendHead
    
    CopyMemory head, Dxb(0), Len(head)
    User = head.UserMsg
    User = Left(User, InStr(User, Chr(0)) - 1)
    head.UserMsg = Winsock1(Index).Tag
    head.UserLen = uLen(Winsock1(Index).Tag)
    Target = GetTargetIndex(User)
    
    If Target < 0 Then Exit Function

    PicSend(Index).Target = Target
    PicSend(Index).Start = True

    PicSend(Index).MaxLen = head.MaxLenth
    PicSend(Index).RevLen = UBound(Dxb) + 1 - Len(head)
    
    CopyMemory Dxb(0), head, Len(head)
    Winsock1(PicSend(Index).Target).SendData Dxb
    Debug.Print PicSend(Index).RevLen, PicSend(Index).MaxLen, 1111
    If PicSend(Index).RevLen >= PicSend(Index).MaxLen Then
        PicSend(Index).Start = False
    End If
End Function
Function GetPicTarget(Index As Integer, Dxb() As Byte) As String
    Dim User As String
    Dim Target As Integer
    Dim head As SendHead
    Dim I As Long
    
    CopyMemory head, Dxb(0), Len(head)
    User = head.UserMsg
    User = Left(User, InStr(User, Chr(0)) - 1)
    
    head.UserMsg = Winsock1(Index).Tag
    head.UserLen = uLen(Winsock1(Index).Tag)


    PicSend(Index).Target = User
    PicSend(Index).Start = True

    PicSend(Index).MaxLen = head.MaxLenth
    PicSend(Index).RevLen = UBound(Dxb) + 1 - Len(head)
    
    PicSend(Index).ByteAddr = VirtualAlloc(ByVal 0&, PicSend(Index).MaxLen + Len(head), MEM_COMMIT, PAGE_READWRITE)
    CopyMemory Dxb(0), head, Len(head)
    CopyMemory ByVal PicSend(Index).ByteAddr, Dxb(0), UBound(Dxb) + 1
    
    PicSend(Index).ByteAddr = PicSend(Index).ByteAddr + Len(head)
    
    If PicSend(Index).RevLen >= PicSend(Index).MaxLen Then
        PicSend(Index).Start = False
        
        Target = GetTargetIndex(User)
        If Target >= 0 Then
            Winsock1(Target).SendData Dxb
        Else
            'GetData User, I
            'SetUserData CStr(I), Dxb
        End If
        VirtualFree ByVal PicSend(Index).ByteAddr, 0, MEM_RELEASE
    End If
End Function
Function PicSendUser(Index As Integer, Dxb() As Byte) As Long
    Dim Addr As Long, Sxb() As Byte, pHead As SendHead, LenS As Long, Target As Integer, I As Long
    
    Addr = PicSend(Index).ByteAddr + PicSend(Index).RevLen
    CopyMemory ByVal Addr, Dxb(0), UBound(Dxb) + 1
    PicSend(Index).RevLen = PicSend(Index).RevLen + UBound(Dxb) + 1

    If PicSend(Index).RevLen >= PicSend(Index).MaxLen Then
        PicSend(Index).Start = False
        
        LenS = Len(pHead) + PicSend(Index).MaxLen
        ReDim Sxb(LenS - 1)
        PicSend(Index).ByteAddr = PicSend(Index).ByteAddr - Len(pHead)
        CopyMemory Sxb(0), ByVal PicSend(Index).ByteAddr, LenS
        
        Target = GetTargetIndex(PicSend(Index).Target)
        If Target >= 0 Then
            Winsock1(Target).SendData Sxb
        Else
            GetData PicSend(Index).Target, I
            SetUserData CStr(I), Sxb
        End If
        VirtualFree ByVal PicSend(Index).ByteAddr, 0, MEM_RELEASE
        
        If PicSend(Index).RevLen > PicSend(Index).MaxLen Then PicSendUser = PicSend(Index).RevLen - PicSend(Index).MaxLen
    End If
End Function
Function LoginMsg(User As String) As String
    Dim Fr As String, Nick As String, Gold As Long, Icon As Long, numID As Long
    Dim sNick As String, sIcon As Long, sNum As Long
    GetData User, numID, , Nick, , , Gold, Icon
    Dim I As Long, tmp() As String
    tmp = GetFriend(User)
    If SafeArrayGetDim(tmp) > 0 Then
        For I = 0 To UBound(tmp)
            GetData tmp(I), sNum, , sNick, , , , sIcon
            LoginMsg = LoginMsg & tmp(I) & "," & sNick & "," & sNum & "," & sIcon & IIf(I < UBound(tmp), "|", "")
        Next
    End If
    LoginMsg = LoginMsg & "-" & Nick & "-" & numID & "-" & Gold & "-" & Icon
End Function

Sub GetData(Name As String, Optional Num As Long, Optional Paswd As String, Optional Nick As String, Optional Tname As String, _
            Optional CardID As String, Optional Gold As Long, Optional Icon As Long)
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from userdata where ID=""" & Name & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    If Not IsMissing(Num) Then Num = Rs("num")
    If Not IsMissing(Paswd) Then Paswd = Rs("password")
    If Not IsMissing(Nick) Then Nick = Rs("nick")
    If Not IsMissing(Tname) Then Tname = Rs("TrueName")
    If Not IsMissing(CardID) Then CardID = Rs("CardID")
    If Not IsMissing(Gold) Then Gold = Rs("Gold")
    If Not IsMissing(Icon) Then Icon = Rs("Icon")
    Rs.Close
    Set Rs = Nothing
End Sub
Sub GetDataFromID(ID As String, Optional Name As String, Optional Paswd As String, Optional Nick As String, Optional Tname As String, _
            Optional CardID As String, Optional Gold As Long, Optional Icon As Long)
    Dim Rs As New ADODB.Recordset
    Rs.Open "Select * from userdata where num=""" & ID & """", Conn, adOpenDynamic, 3
    If Rs.EOF Then
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    If Not IsMissing(Name) Then Name = Rs("ID")
    If Not IsMissing(Paswd) Then Paswd = Rs("password")
    If Not IsMissing(Nick) Then Nick = Rs("nick")
    If Not IsMissing(Tname) Then Tname = Rs("TrueName")
    If Not IsMissing(CardID) Then CardID = Rs("CardID")
    If Not IsMissing(Gold) Then Gold = Rs("Gold")
    If Not IsMissing(Icon) Then Icon = Rs("Icon")
    Rs.Close
    Set Rs = Nothing
End Sub
Sub UserToUser(Index As Integer, Dxb() As Byte)
    Dim I As Long, User As String
    Dim head As msgHead
    CopyMemory head, Dxb(0), Len(head)
    User = head.UserMsg
    User = Left(User, InStr(User, Chr(0)) - 1)
    head.UserMsg = Winsock1(Index).Tag
    head.UserLen = uLen(Winsock1(Index).Tag)
    CopyMemory Dxb(0), head, Len(head)
    I = GetTargetIndex(User)
    If I >= 0 Then
        Winsock1(I).SendData Dxb
    Else
        'GetData User, I
        'SetUserData CStr(I), Dxb
    End If
End Sub

Function GetTargetIndex(User As String) As Integer
    For GetTargetIndex = 0 To UBound(ConnectState)
        If Winsock1(GetTargetIndex).Tag = User And Winsock1(GetTargetIndex).State = sckConnected Then Exit Function
    Next
    GetTargetIndex = -1
End Function
Sub MakeDir(Uname As String)
    On Error GoTo ProErr
    MakeDirBase
    MkDir App.Path & "\Data\" & Uname
    Exit Sub
ProErr:
End Sub
Sub MakeDirBase()
    On Error GoTo ProErr
    MkDir App.Path & "\Data"
    Exit Sub
ProErr:
End Sub

'寻找空闲的sock
Public Function FindFreeSocket()
    Dim SockCount, I As Integer
    SockCount = UBound(ConnectState)
    For I = 0 To SockCount
        If ConnectState(I) = FREE Then
            FindFreeSocket = I
            Exit Function
        End If
    Next I
    ReDim Preserve ConnectState(0 To SockCount + 1)
    ReDim Preserve Connected(0 To SockCount + 1)
    FindFreeSocket = UBound(ConnectState)
End Function

Public Sub Dtxt(ParamArray Texts())
    Dim text As String, I As Long
    If Len(Text1.text) > 30000 Then Text1.text = ""
    For I = 0 To UBound(Texts)
        text = text & Texts(I) & ", "
    Next
    Debug.Print text
    Text1.text = text & vbCrLf & Text1.text
End Sub

