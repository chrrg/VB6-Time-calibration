VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "北京时间校准"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5850
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "校准时间"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   2280
   End
   Begin VB.Label Label6 
      Caption         =   "软件作者：CH"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label Label4 
      Caption         =   "北京时间："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "本机："
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WinHttp As Object
Dim api As String
Dim cha As Long
Dim time As Long
Dim oldtick As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Private CurDate As SYSTEMTIME
Private Function getTime()
On Error Resume Next
    WinHttp.Open "GET", api, True
    WinHttp.send
    WinHttp.WaitForResponse
    If InStr(1, WinHttp.ResponseText, """t"":""") = 0 Then Exit Function
    getTime = Split(Split(WinHttp.ResponseText, """t"":""")(1), """")(0)
End Function
Private Function setTime(t, Optional Milliseconds = 0) '时间戳
    Label2.Caption = FromUnixTime(t, 8)
    DoEvents
    Dim lpSystemTime As SYSTEMTIME, succ As Long
    Dim nowtime As String
    nowtime = FromUnixTime(t, 0) '不能填8
    With CurDate
        .wYear = Year(nowtime)
        .wMonth = Month(nowtime)
        .wDay = Day(nowtime)
        .wHour = Hour(nowtime)
        .wMinute = Minute(nowtime)
        .wSecond = Second(nowtime)
        .wMilliseconds = Milliseconds
    End With
    succ = SetSystemTime(CurDate)
    Timer2.Enabled = False
    If Milliseconds <> 0 Then
        Delay 1000 - Milliseconds
        time = t + 1
    Else
        time = t
    End If
    oldtick = GetTickCount
    Timer2.Enabled = True
    If succ = 0 Then
        Label5.ForeColor = RGB(200, 20, 10)
        Label5.Caption = "修改时间失败，请尝试用管理员打开！"
    Else
        Label5.ForeColor = RGB(10, 200, 20)
        Label5.Caption = "恭喜！自动校准时间成功！"
    End If
    
End Function
Private Sub Command1_Click()
'方法零：直接法
Command1.Enabled = False
Label2.Caption = ""
Label5.Caption = ""
Timer2.Enabled = False
DoEvents
Dim ti As String
ti = getTime
If Not IsNumeric(ti) Or ti = "0" Then
    MsgBox "校准失败！返回：" & ti
    Command1.Enabled = True
    Exit Sub
End If
Dim t As Double
t = getTime / 1000
'Sleep Int(1000 - (t - Int(t)) * 1000)
'setTime Int(t) + 1
setTime Int(t), Int((t - Int(t)) * 1000)
DoEvents
Command1.Enabled = True
'sleep 1000 - t Mod 1000

'方法一：循环
'Label2.Caption = ""
'Timer2.Enabled = False
'Dim oldt As String, newt As String
'oldt = getTime
'If Not IsNumeric(oldt) Then
'    MsgBox "校准失败！返回：" & oldt
'    Exit Sub
'End If
'Dim i As Long
'Do
'newt = getTime
'i = i + 1
'If oldt <> newt Then
'    setTime newt
'    Exit Do
'End If
'DoEvents
'Loop Until 0
'Exit Sub
'方法二：二分法
'cha = 500 '1秒
'Dim i As Long, j As Long
'Dim allSleep As Long
'i = getTime
'Sleep cha
'allSleep = allSleep + cha
'cha = cha / 2
'j = getTime


'MsgBox FromUnixTime(getTime, 8)
'If IsNumeric(xml.ResponseText) Then
'End If
End Sub

Private Sub Form_Load()
On Error GoTo errs:
    Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttp.SetTimeouts 1000, 1000, 1000, 1000 '设置操作超时时间
    WinHttp.Option(4) = 13056 '忽略错误标志
    WinHttp.Option(6) = True '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。
    'api = "http://47.115.38.15/tbapi/ch/tbTime.php" '私人淘宝时间api地址不推荐
    api = "http://api.m.taobao.com/rest/api3.do?api=mtop.common.getTimestamp"
    Exit Sub
errs:
MsgBox "启动发生错误" & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label6_Click()
    MsgBox "本软件通过HTTP协议调用淘宝时间api进行时间校准，仅供学习使用！——CH", vbOKOnly, "彩蛋"
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Now
End Sub

'参数：strTime:要转换的时间；intTimeZone：该时间对应的时区
'返回值：strTime相对于1970年1月1日午夜0点经过的秒数
'示例：ToUnixTime("2008-5-23 10:51:0", +8)，返回值为1211511060
Function ToUnixTime(strTime, intTimeZone)
    If IsEmpty(strTime) Or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
     ToUnixTime = DateAdd("h", -intTimeZone, strTime)
     ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
End Function
         
'把UNIX时间戳转换为标准时间
'参数：intTime:要转换的UNIX时间戳；intTimeZone：该时间戳对应的时区
'返回值：intTime所代表的标准时间
'示例：FromUnixTime("1211511060", +8)，返回值2008-5-23 10:51:0
Function FromUnixTime(intTime, intTimeZone)
    If IsEmpty(intTime) Or Not IsNumeric(intTime) Then
        FromUnixTime = Now()
        Exit Function
    End If
    If IsEmpty(intTime) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
    FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0")
    FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)
End Function

Private Sub Timer2_Timer()
Label2.Caption = FromUnixTime(time + (GetTickCount - oldtick) / 1000, 8)
End Sub
Public Sub Delay(Msec As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + Msec
    Do
        Sleep 1
        DoEvents
    Loop While GetTickCount < EndTime
End Sub
