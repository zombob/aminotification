Attribute VB_Name = "Tray"
'* Copyright (c) 2010 著作权由张鑫所有。著作权人保留一切权利。
'*
'* 这份授权条款，在使用者符合以下三条件的情形下，授予使用者使用及再散播本
'* 软件包装原始码及二进位可执行形式的权利，无论此包装是否经改作皆然：
'*
'* * 对于本软件源代码的再散播，必须保留上述的版权宣告、此三条件表列，以
'*   及下述的免责声明。
'* * 对于本套件二进位可执行形式的再散播，必须连带以文件以及／或者其他附
'*   于散播包装中的媒介方式，重制上述之版权宣告、此三条件表列，以及下述
'*   的免责声明。
'* * 未获事前取得书面许可，不得使用张鑫或本软件贡献者之名称，
'*   来为本软件之衍生物做任何表示支持、认可或推广、促销之行为。
'*
'* 免责声明：本软件是由张鑫及本软件之贡献者以现状（"as is"）提供，本软件
'* 包装不负任何明示或默示之担保责任，包括但不限于就适售性以及特定目的的
'* 适用性为默示性担保。张鑫及本软件之贡献者，无论任何条件、无论成因或任
'* 何责任主义、无论此责任为因合约关系、无过失责任主义或因非违约之侵权
'* （包括过失或其他原因等）而起，对于任何因使用本软件包装所产生的任何直
'* 接性、间接性、偶发性、特殊性、惩罚性或任何结果的损害（包括但不限于替
'* 代商品或劳务之购用、使用损失、资料损失、利益损失、业务中断等等），不
'* 负任何责任，即在该种使用已获事前告知可能会造成此类损害的情形下亦然。

'* Copyright (c) 2010, ZhangXin. All rights reserved.
'*
'* Redistribution and use in source and binary forms, with or without
'* modification, are permitted provided that the following conditions are met:
'*
'*     * Redistributions of source code must retain the above copyright
'*       notice, this list of conditions and the following disclaimer.
'*     * Redistributions in binary form must reproduce the above copyright
'*       notice, this list of conditions and the following disclaimer in the
'*       documentation and/or other materials provided with the distribution.
'*     * Neither the name of the ZhangXin nor the names of its contributors
'*       may be used to endorse or promote products derived from this software
'*       without specific prior written permission.
'*
'* THIS SOFTWARE IS PROVIDED BY ZHANGXIN AND CONTRIBUTORS "AS IS" AND ANY
'* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
'* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'* DISCLAIMED. IN NO EVENT SHALL ZHANGXIN AND CONTRIBUTORS BE LIABLE FOR ANY
'* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
'* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
'* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.



Option Explicit

'使用高分辨率图标所用的API
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const IMAGE_ICON = 1
'系统托盘
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
Private Const WM_USER As Long = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Private Const NOTIFYICON_VERSION = 3
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Private Const WM_CLOSE As Long = &H10
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

Public Enum bFlag
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10 '关闭提示音标志
End Enum

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

'鼠标事件
Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public ni As NOTIFYICONDATA
Public Sub TrayAddIcon(ByVal MyForm As Form, ByVal MyIcon As String, ByVal ToolTip As String, Optional ByVal bFlag As bFlag)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = LoadImage(App.hInstance, MyIcon, IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS)
        .szTip = ToolTip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, ni)

End Sub

Public Sub TrayRemoveIcon()

    Shell_NotifyIcon NIM_DELETE, ni
    
End Sub

Public Sub TrayBalloon(ByVal MyForm As Form, ByVal sBaloonText As String, sBallonTitle As String, Optional ByVal bFlag As bFlag)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_INFO
        .dwInfoFlags = bFlag
        .szInfoTitle = sBallonTitle & vbNullChar
        .szInfo = sBaloonText & vbNullChar
    End With

    Shell_NotifyIcon NIM_MODIFY, ni

End Sub

Public Sub TrayTip(ByVal MyForm As Form, ByVal sTipText As String)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = sTipText & vbNullChar
    End With

    Shell_NotifyIcon NIM_MODIFY, ni

End Sub

