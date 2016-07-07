VERSION 5.00
Begin VB.Form RealTimeForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "实时车次售票状况显示设置"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "realtimeoptine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox PAText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox UserNameText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox RowText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   5
      Top             =   225
      Width           =   735
   End
   Begin VB.TextBox IntervalText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "密码 ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "数据库用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1365
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "上卷行数 ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "刷新间隔时间 ："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   795
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2895
      TabIndex        =   2
      Top             =   285
      Width           =   450
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      TabIndex        =   0
      Top             =   765
      Width           =   375
   End
End
Attribute VB_Name = "RealTimeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub TestKey(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    Beep
End If
End Sub
Public Sub SaveRealTimeOption()
Dim filenum As Integer
filenum = FreeFile
Open CurrentPath + "database.cfg" For Output As #filenum
Print #filenum, RowText.Text
Print #filenum, IntervalText.Text
Print #filenum, UserNameText.Text
Print #filenum, PAText.Text
Close #filenum
iScrollRowofRealTime = Val(RowText.Text)
iIntervaltimeofRealTime = Val(IntervalText.Text)
If flagofdatabaseform = True Then
   DatabaseForm.DisTimer.Interval = iIntervaltimeofRealTime * 1000
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
     SaveRealTimeOption
     Unload RealTimeForm
Case 1
     Unload RealTimeForm
End Select
End Sub

Private Sub Form_Load()
Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant
filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "database.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Input #filenum, c
Input #filenum, d
Close #filenum
RowText.Text = a
IntervalText.Text = b
UserNameText.Text = c
PAText.Text = d
RealTimeForm.Top = gtop + gheight
RealTimeForm.Left = 2500

Exit Sub
dealerror:
Close #filenum
RowText.Text = "5"
IntervalText.Text = "5"
UserNameText.Text = "Guest"
PAText.Text = "Guest"

'sUserName = "Guest"
'sPassword = "Guest"
iScrollRowofRealTime = 5
iIntervaltimeofRealTime = 5
RealTimeForm.Top = gtop + gheight
RealTimeForm.Left = 2500
End Sub




Private Sub IntervalText_LostFocus()
If Val(IntervalText.Text) > 60 Or Val(IntervalText.Text) < 1 Then
       action = MsgBox("刷新间隔时间范围为(1，60)秒!", vbInformation, "消息")
       IntervalText.Text = "5"
End If
End Sub

Private Sub RowText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub IntervalText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub RowText_LostFocus()
If Val(RowText.Text) > 20 Or Val(RowText.Text) < 1 Then
       action = MsgBox("上卷行数范围为(1，20)行!", vbInformation, "消息")
       RowText.Text = "10"
End If
End Sub

