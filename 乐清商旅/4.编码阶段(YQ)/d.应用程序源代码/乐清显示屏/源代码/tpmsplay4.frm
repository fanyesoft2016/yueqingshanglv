VERSION 5.00
Begin VB.Form CoordForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "播放窗口映射"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tpmsplay4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5.848
   ScaleMode       =   0  'User
   ScaleWidth      =   6.562
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckAuto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "进入程序自动播放前一次内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   285
      TabIndex        =   11
      Top             =   2280
      Width           =   3390
   End
   Begin VB.CommandButton Command2 
      Caption         =   "映射窗口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2490
      TabIndex        =   10
      Top             =   2835
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
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
      Left            =   1305
      TabIndex        =   9
      Top             =   2820
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Left            =   60
      TabIndex        =   8
      Top             =   2820
      Width           =   975
   End
   Begin VB.TextBox HeightText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox WidthText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2025
      TabIndex        =   6
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox YText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2025
      TabIndex        =   5
      Top             =   705
      Width           =   1455
   End
   Begin VB.TextBox XText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   1
      Top             =   195
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "映射窗口高度 ："
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
      Index           =   3
      Left            =   165
      TabIndex        =   4
      Top             =   1815
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "左上角 Y ："
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
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "映射窗口宽度 ："
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
      Left            =   195
      TabIndex        =   2
      Top             =   1290
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "左上角 X ："
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
      Left            =   285
      TabIndex        =   0
      Top             =   210
      Width           =   1335
   End
End
Attribute VB_Name = "CoordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub TestKey(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    'action = MsgBox("valid ", 16, "information")
    Beep
End If
End Sub
Private Sub SaveCoord()
filenum% = FreeFile
'Open "coord.cfg" For Input As #filenum
'Close #filenum
Open CurrentPath + "coord.cfg" For Output As #filenum
Print #filenum, XText.Text
Print #filenum, YText.Text
Print #filenum, WidthText.Text
Print #filenum, HeightText.Text
Print #filenum, CheckAuto.Value
Close #filenum
gleft = Val(XText.Text) * TwipspercentPixelX
gtop = Val(YText.Text) * TwipspercentPixelY
gwidth = Val(WidthText.Text) * TwipspercentPixelX
gheight = Val(HeightText.Text) * TwipspercentPixelY
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   SaveCoord
   Unload CoordForm
ElseIf Index = 1 Then
   Unload CoordForm
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "映射窗口" Then
 If XText.Text = "" Or YText.Text = "" Or WidthText.Text = "" Or HeightText.Text = "" Then
      action = MsgBox("坐标值未输入完整！", 16, "错误")
      Exit Sub
 End If
 gleft = Val(XText.Text) * TwipspercentPixelX
 gtop = Val(YText.Text) * TwipspercentPixelY
 gwidth = Val(WidthText.Text) * TwipspercentPixelX
 gheight = Val(HeightText.Text) * TwipspercentPixelY
 Command2.Caption = "关闭窗口"
 Call AppForm.DisplayWindow
Else
 Command2.Caption = "映射窗口"
 Unload DisForm
End If
End Sub

Private Sub Form_Load()

Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant
Dim e As Variant
filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "coord.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Input #filenum, c
Input #filenum, d
Input #filenum, e
Close #filenum
XText.Text = a
YText.Text = b
WidthText.Text = c
HeightText.Text = d
CheckAuto.Value = e
FlagofAuto = CheckAuto.Value
CoordForm.Top = gtop + gheight
CoordForm.Left = 2500

Exit Sub
dealerror:
XText.Text = "0"
YText.Text = "0"
WidthText.Text = "400"
HeightText.Text = "300"
CheckAuto.Value = 0
gleft = 0
gtop = 0
gwidth = 400 * TwipspercentPixelX
gheight = 300 * TwipspercentPixelY
FlagofAuto = 0
Close #filenum
CoordForm.Top = gtop + gheight
CoordForm.Left = 2500
End Sub

Private Sub HeightText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub HeightText_LostFocus()
If Val(HeightText.Text) < 0 Or Val(HeightText.Text) > 1280 Then
    action = MsgBox("输入值应在(0，1280)!", vbInformation, "消息")
    HeightText.Text = ""
End If

End Sub

Private Sub WidthText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub WidthText_LostFocus()
If Val(WidthText.Text) < 0 Or Val(WidthText.Text) > 2560 Then
    action = MsgBox("输入值应在(0，2560)!", vbInformation, "消息")
    WidthText.Text = ""
End If

End Sub

Private Sub XText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii

End Sub

Private Sub XText_LostFocus()
If Val(XText.Text) < 0 Or Val(XText.Text) > 1280 Then
    action = MsgBox("输入值应在(0，1280)!", vbInformation, "消息")
    XText.Text = ""
End If
End Sub

Private Sub YText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub YText_LostFocus()
If Val(YText.Text) < 0 Or Val(YText.Text) > 1280 Then
    action = MsgBox("输入值应在(0，1280)!", vbInformation, "消息")
    YText.Text = ""
End If
End Sub
