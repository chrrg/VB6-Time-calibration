VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����ʱ��У׼"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5850
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "У׼ʱ��"
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
      Caption         =   "������ߣ�CH"
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
      Caption         =   "����ʱ�䣺"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "������"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
Private Function setTime(t, Optional Milliseconds = 0) 'ʱ���
    Label2.Caption = FromUnixTime(t, 8)
    DoEvents
    Dim lpSystemTime As SYSTEMTIME, succ As Long
    Dim nowtime As String
    nowtime = FromUnixTime(t, 0) '������8
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
        Label5.Caption = "�޸�ʱ��ʧ�ܣ��볢���ù���Ա�򿪣�"
    Else
        Label5.ForeColor = RGB(10, 200, 20)
        Label5.Caption = "��ϲ���Զ�У׼ʱ��ɹ���"
    End If
    
End Function
Private Sub Command1_Click()
'�����㣺ֱ�ӷ�
Command1.Enabled = False
Label2.Caption = ""
Label5.Caption = ""
Timer2.Enabled = False
DoEvents
Dim ti As String
ti = getTime
If Not IsNumeric(ti) Or ti = "0" Then
    MsgBox "У׼ʧ�ܣ����أ�" & ti
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

'����һ��ѭ��
'Label2.Caption = ""
'Timer2.Enabled = False
'Dim oldt As String, newt As String
'oldt = getTime
'If Not IsNumeric(oldt) Then
'    MsgBox "У׼ʧ�ܣ����أ�" & oldt
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
'�����������ַ�
'cha = 500 '1��
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
    WinHttp.SetTimeouts 1000, 1000, 1000, 1000 '���ò�����ʱʱ��
    WinHttp.Option(4) = 13056 '���Դ����־
    WinHttp.Option(6) = True 'Ϊ True ʱ��������ҳ���ض�����תʱ�Զ���ת��False ���Զ���ת����ȡ����˷��ص�302״̬��
    'api = "http://47.115.38.15/tbapi/ch/tbTime.php" '˽���Ա�ʱ��api��ַ���Ƽ�
    api = "http://api.m.taobao.com/rest/api3.do?api=mtop.common.getTimestamp"
    Exit Sub
errs:
MsgBox "������������" & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label6_Click()
    MsgBox "�����ͨ��HTTPЭ������Ա�ʱ��api����ʱ��У׼������ѧϰʹ�ã�����CH", vbOKOnly, "�ʵ�"
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Now
End Sub

'������strTime:Ҫת����ʱ�䣻intTimeZone����ʱ���Ӧ��ʱ��
'����ֵ��strTime�����1970��1��1����ҹ0�㾭��������
'ʾ����ToUnixTime("2008-5-23 10:51:0", +8)������ֵΪ1211511060
Function ToUnixTime(strTime, intTimeZone)
    If IsEmpty(strTime) Or Not IsDate(strTime) Then strTime = Now
    If IsEmpty(intTimeZone) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0
     ToUnixTime = DateAdd("h", -intTimeZone, strTime)
     ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
End Function
         
'��UNIXʱ���ת��Ϊ��׼ʱ��
'������intTime:Ҫת����UNIXʱ�����intTimeZone����ʱ�����Ӧ��ʱ��
'����ֵ��intTime������ı�׼ʱ��
'ʾ����FromUnixTime("1211511060", +8)������ֵ2008-5-23 10:51:0
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
