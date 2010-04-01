VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form notification 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "AMI通知"
   ClientHeight    =   4980
   ClientLeft      =   8100
   ClientTop       =   4185
   ClientWidth     =   3585
   Icon            =   "notification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   3585
   Begin VB.Frame otherSetFrame 
      BackColor       =   &H80000004&
      Caption         =   "其他设置"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   3375
      Begin VB.CheckBox BalloonCheck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "气泡提示"
         ForeColor       =   &H80000008&
         Height          =   185
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox TrackCheck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "打开监控框"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame PopSetFrame 
      BackColor       =   &H80000004&
      Caption         =   "弹屏设置"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox ExtenTXT 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "若要弹出所有分机来电,请清空此文本框"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox PopAddrTXT 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Text            =   "http://"
         ToolTipText     =   "电话号码用%s表示,否则号码将自动加到末尾"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "分机,多个分机用逗号隔开"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2070
      End
   End
   Begin VB.Frame HostSetFrame 
      BackColor       =   &H80000004&
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3375
      Begin VB.TextBox NameTXT 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox PortTXT 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox PswTXT 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox HostTXT 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "登录名"
         Height          =   180
         Left            =   2400
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "密码"
         Height          =   180
         Left            =   2400
         TabIndex        =   13
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "端口"
         Height          =   180
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "服务器地址"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton loginCommand 
      Caption         =   "登录"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox TrackTXT 
      Height          =   4695
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mSysPopup 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu mShow 
         Caption         =   "显示"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''尚未完成的部分:
'''''''''''''''1.一个心跳程序，检测与服务器连接的状态
'''''''''''''''2.改进登录方式（使用加密密码登录方式）
'''''''''''''''3.把密码加密存贮至注册表
'''''''''''''''4.托盘图标的真彩色

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const sckClosed = 0             '缺省的。关闭
'Private Const sckOpen = 1               '打开
'Private Const sckListening = 2          '侦听
'Private Const sckConnectionPending = 3  '连接挂起
'Private Const sckResolvingHost = 4      '识别主机
'Private Const sckHostResolved = 5       '已识别主机
'Private Const sckConnecting = 6         '正在连接
Private Const sckConnected = 7           '已连接
'Private Const sckClosing = 8            '同级人员正在关闭连接
'Private Const sckError = 9              '错误
Private Logined As Boolean
Dim LoginTXT As String

Dim WSHShell As Object ''''WScript.Shell用于读写注册表
Const regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\AMI Notification\" '写入注册表的路径

Private Sub BalloonCheck_Click()
WriteReg BalloonCheck
End Sub

Private Sub ExtenTXT_Change()
ExtenTXT.Text = Trim(ExtenTXT.Text)
WriteReg ExtenTXT
End Sub


Private Sub Form_Load()
    '载入系统托盘
    TrayAddIcon notification, App.Path & "\ico.ico", "来电话时,我会喊哦!"
    
    ''添加一个WScript.Shell
    Set WSHShell = CreateObject("WScript.Shell")
    Logined = False
    
    HostTXT.Text = ReadReg(HostTXT)
    PortTXT.Text = ReadReg(PortTXT)
    NameTXT.Text = ReadReg(NameTXT)
    PswTXT.Text = ReadReg(PswTXT)
    TrackCheck.Value = ReadReg(TrackCheck)
    BalloonCheck.Value = ReadReg(BalloonCheck)
    
    If ReadReg(PopAddrTXT) <> "" Then
        PopAddrTXT.Text = ReadReg(PopAddrTXT)
    Else
        PopAddrTXT.Text = "http://"
    End If
    ExtenTXT.Text = ReadReg(ExtenTXT)
    

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '气泡单击时的鼠标事件
    Dim Result As Long
    Dim cEvent As Single
    cEvent = x / Screen.TwipsPerPixelX

    Select Case cEvent

    Case MouseMove
        'Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "左键放开"
    Case LeftDown
        Debug.Print "左键按下"
        notification.WindowState = 0
        notification.Show
    Case LeftDbClick
        Debug.Print "左键双击"
    Case MiddleUp
        Debug.Print "中间键放开"
    Case MiddleDown
        Debug.Print "中间键按下"
    Case MiddleDbClick
        Debug.Print "中间键单击"
    Case RightUp
        Debug.Print "右健放开"
    Case RightDown
        Debug.Print "右健按下"
        '单击后移出
        Result = SetForegroundWindow(Me.hwnd)
        '当时显示
        Me.PopupMenu Me.mSysPopup
    Case RightDbClick
        Debug.Print "右健双击"
    Case BalloonClick
        Debug.Print "单击气泡"

    End Select

End Sub


Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Shell_NotifyIcon NIM_DELETE, Tray
TrayRemoveIcon
End Sub

Private Sub HostTXT_Change()
'HostTXT.Text = Trim(HostTXT.Text) '去空格
If ReadReg(HostTXT) <> HostTXT.Text Then
    WriteReg HostTXT
    
    WinsockClient.Close
    If WinsockClient.State = sckClosed Then
        Me.Caption = "AMI通知"
    End If
    WinsockClient.RemoteHost = HostTXT.Text
End If
End Sub

Private Sub PortTXT_Change()
PortTXT.Text = Trim(PortTXT.Text)
If ReadReg(PortTXT) <> PortTXT.Text Then
    WriteReg PortTXT

    WinsockClient.Close
    If WinsockClient.State = sckClosed Then
        Me.Caption = "AMI通知"
    End If
    If PortTXT.Text > 65535 Then
        PortTXT.Text = 65534
    End If
    If PortTXT.Text <> "" Then
        WinsockClient.RemotePort = PortTXT.Text
    End If
End If
End Sub


Private Sub NameTXT_Change()
If ReadReg(NameTXT) <> NameTXT.Text Then
    WriteReg NameTXT
End If
End Sub

Private Sub PopAddrTXT_Change()
If ReadReg(PopAddrTXT) <> PopAddrTXT.Text Then
    WriteReg PopAddrTXT
End If
End Sub


Private Sub loginCommand_Click()
    
    If WinsockClient.State = sckConnected Then
        WinsockClient.Close
        Logined = False
        WinsockClient_Closed
        Me.Caption = "AMI通知-已断开"
        loginCommand.Caption = "重新连接"
    Else
        LoginTXT = "Action: login" & vbCrLf & "Username: " & NameTXT.Text & vbCrLf & "Secret: " & PswTXT.Text
        If HostTXT.Text <> "" And PortTXT.Text <> "" Then
            WinsockClient.Close
            WinsockClient.RemoteHost = HostTXT.Text
            WinsockClient.RemotePort = PortTXT.Text
            WinsockClient.Connect
        End If
    End If
End Sub

Private Sub loginCommand_GotFocus() '''''当按钮得到焦点时，生成发送用的文本logintxt，可一次性防止用户名密码改变

End Sub

Private Sub PortTXT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 'backspace
            Exit Sub
        Case 48 To 57 '0-9
            Exit Sub
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub PswTXT_Change()
WriteReg PswTXT
End Sub

Private Sub TrackCheck_Click()
    If TrackCheck.Value = 1 Then
        TrackTXT.Visible = True
        Me.Width = 8445
    ElseIf TrackCheck.Value = 0 Then
        TrackTXT.Visible = False
        Me.Width = 3705
    End If
    WriteReg TrackCheck
End Sub


Private Sub WinsockClient_Close()
WinsockClient_Closed
End Sub

Private Sub WinsockClient_Connect()
notification.Caption = "已发现服务器"
WinsockClient.SendData (LoginTXT & vbCrLf & vbCrLf)
End Sub

Private Sub winsockclient_DataArrival(ByVal bytesTotal As Long)

    Dim arrival() As Byte, i As Long, s As String
    ReDim arrival(bytesTotal - 1) As Byte
    Dim arrivalStr As String
    'Dim ccc
    
    WinsockClient.GetData arrival, vbByte, bytesTotal
    For i = 0 To bytesTotal - 1
        arrivalStr = arrivalStr & Chr$(arrival(i))
        'Debug.Print i&; ":" & Chr$(arrival(i))
    Next
    If Logined = True Then
        PopDialEvent (arrivalStr)
    Else
        WaitLogin (arrivalStr)
        'trackTXT.Text = trackTXT.Text & arrivalStr ''''''''''''''''把trackTXT....放到这里，当登录成功后不再显示内容
    End If
        
    ''''''''限定trackTXT的文本大小
    If Len(TrackTXT.Text) > 300 Then
        TrackTXT.Text = ""
    End If
    TrackTXT.Text = TrackTXT.Text & arrivalStr
End Sub

Private Function PopDialEvent(Str As String) '把收到的string，按照两个换行符分段，并决定是否弹屏
    Dim Events() As String
    Events = Split(Str, vbCrLf & vbCrLf)
    
    For i = 0 To UBound(Events)
        
        If FoundPopEvent(Events(i)) Then
            Popup (Events(i))
        End If
    Next
End Function

Private Function FoundPopEvent(Str As String) As Boolean '使用条件过滤Event,根据具体需求更改###################################################
    'If InStr(str, "Event: Dial") And InStr(str, "Source: DAHDI") And InStr(str, "Destination: SIP") Then
    If Mid(Str, 1, 11) = "Event: Dial" And InStr(Str, "Source: ") And InStr(Str, "Destination: SIP") Then
        FoundPopEvent = True
    Else
        FoundPopEvent = False
    End If
End Function

Private Function Popup(Str As String)   '实施弹屏
    Dim DestPhone, CallerID, PopAddrStr As String
    '''获取目的电话号码
    StartNum = InStr(Str, "Destination: ")               '先找到“Destination: ”
    StartNum = InStr(StartNum, Str, "/") + 1             '再找此行中的“/”
    StopNum = InStr(StartNum, Str, "-")   '找到这一行中的“-”
    DestPhone = Mid(Str, StartNum, StopNum - StartNum)   '获取“/”“-”中间的部分
    '''''''获取callerid
    StartNum = InStr(Str, "CallerID: ") + 10
    StopNum = InStr(StartNum, Str, vbCrLf)
    CallerID = Mid(Str, StartNum, StopNum - StartNum)
    
    If InStr(PopAddrTXT.Text, "%s") Then
        PopAddrStr = Replace(PopAddrTXT.Text, "%s", CallerID)
    Else
        PopAddrStr = PopAddrTXT.Text & CallerID
    End If
    
    If Trim(ExtenTXT.Text) = "" Or InStr(ExtenTXT.Text, DestPhone) Then '当分机文本框为空时,弹出所有
        '气泡显示
        If BalloonCheck.Value = 1 Then
            TrayBalloon notification, "是" & vbCrLf & CallerID & "打给" & DestPhone & "的", "来电话啦", NIIF_INFO
        End If
        Shell "cmd /c start " & PopAddrStr '''''使用默认浏览器弹屏
        Debug.Print "到达" & DestPhone
    End If
End Function

Private Function WaitLogin(Str As String)
    If InStr(Str, "Response: Success") Then
        doLoginedSuccess
    ElseIf InStr(Str, "Response: Error") Then
        doLoginedError
    End If
End Function
Private Sub WinsockClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Number & ":" & Description
    HostTXT.Enabled = True
    PortTXT.Enabled = True
    NameTXT.Enabled = True
    PswTXT.Enabled = True
End Sub

Private Function doLoginedSuccess()
    Logined = True
    loginCommand.Caption = "成功登录"
    notification.Caption = "AMI通知-已登录"
    
    'WriteReg HostTXT '登录成功时把服务器数据写入注册表
    'WriteReg PortTXT
    'WriteReg NameTXT
    'WriteReg PswTXT
    
    HostTXT.Enabled = False
    PortTXT.Enabled = False
    NameTXT.Enabled = False
    PswTXT.Enabled = False
End Function

Private Function doLoginedError()
    MsgBox "好好想想登录名和密码", vbExclamation, "登录错误！"
End Function


Private Function WriteReg(Ob As Object)
Dim TN As String
TN = TypeName(Ob)
If TN = "TextBox" Then
    WSHShell.regwrite regPath & Ob.Name, Ob.Text, "REG_SZ"
ElseIf TN = "CheckBox" Then
    WSHShell.regwrite regPath & Ob.Name, Ob.Value, "REG_SZ"
End If
End Function

Private Function ReadReg(Ob As Object) As String
    On Error Resume Next
    ReadReg = WSHShell.RegRead(regPath & Ob.Name)
    'If txt <> "" Then
    '    TB.Text = txt
    'End If
End Function

Private Function WinsockClient_Closed()
    Logined = False
    notification.Caption = "已断开服务器"
    loginCommand.Caption = "登录"

    HostTXT.Enabled = True
    PortTXT.Enabled = True
    NameTXT.Enabled = True
    PswTXT.Enabled = True
End Function

Private Sub mShow_Click()
    
    notification.WindowState = 0
    notification.Show

End Sub

Private Sub mExit_Click()

    Unload Me

End Sub
