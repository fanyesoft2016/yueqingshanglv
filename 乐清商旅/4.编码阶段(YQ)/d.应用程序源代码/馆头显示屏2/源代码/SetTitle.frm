VERSION 5.00
Begin VB.Form frmSetTitle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置标题"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "SetTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5865
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Width           =   4245
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
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1485
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
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "标题 Title:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "左边 Left："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   195
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "顶部 Top："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   195
      Width           =   1155
   End
End
Attribute VB_Name = "frmSetTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   SaveCoord
   Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim a As Variant
Dim b As Variant
Dim c As Variant

filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "Title.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Input #filenum, c

Close #filenum
XText.Text = a
YText.Text = b
txtTitle.Text = c
DatabaseForm.lblTitle.Caption = c

frmSetTitle.Top = gtop + gheight
frmSetTitle.Left = 2500

Exit Sub
dealerror:
XText.Text = DatabaseForm.lblTitle.Left
YText.Text = DatabaseForm.lblTitle.Top
txtTitle.Text = DatabaseForm.lblTitle.Caption
gleft = 0
gtop = 2000
gwidth = 400 * TwipspercentPixelX
gheight = 300 * TwipspercentPixelY
Close #filenum
frmSetTitle.Top = gtop + gheight
frmSetTitle.Left = 2500
End Sub

Private Sub XText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii

End Sub

'Private Sub XText_LostFocus()
'If Val(XText.Text) < 0 Or Val(XText.Text) > 1280 Then
'    action = MsgBox("输入值应在(0，1280)!", vbInformation, "消息")
'    XText.Text = ""
'End If
'End Sub

Private Sub YText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

'Private Sub YText_LostFocus()
'If Val(YText.Text) < 0 Or Val(YText.Text) > 1280 Then
'    action = MsgBox("输入值应在(0，1280)!", vbInformation, "消息")
'    YText.Text = ""
'End If
'End Sub
Private Sub SaveCoord()
filenum% = FreeFile
'Open "coord.cfg" For Input As #filenum
'Close #filenum
Open CurrentPath + "Title.cfg" For Output As #filenum
Print #filenum, XText.Text
Print #filenum, YText.Text
Print #filenum, txtTitle.Text

Close #filenum
DatabaseForm.lblTitle.Left = Val(XText.Text)
DatabaseForm.lblTitle.Top = Val(YText.Text)
DatabaseForm.lblTitle.Caption = txtTitle.Text
End Sub

Public Sub TestKey(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
    'action = MsgBox("valid ", 16, "information")
    Beep
End If
End Sub
