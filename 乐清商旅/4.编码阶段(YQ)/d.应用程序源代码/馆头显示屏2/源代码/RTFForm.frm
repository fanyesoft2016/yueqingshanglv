VERSION 5.00
Begin VB.Form RTFForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RTF文件播放设置"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "RTFForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
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
      Height          =   405
      Index           =   1
      Left            =   1980
      TabIndex        =   5
      Top             =   1545
      Width           =   1125
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
      Height          =   405
      Index           =   0
      Left            =   495
      TabIndex        =   4
      Top             =   1545
      Width           =   1125
   End
   Begin VB.TextBox PText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1815
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox ScrollText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1815
      TabIndex        =   1
      Top             =   915
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "一屏字符数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   300
      TabIndex        =   2
      Top             =   435
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "滚动字符数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   285
      TabIndex        =   0
      Top             =   990
      Width           =   1395
   End
End
Attribute VB_Name = "RTFForm"
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
Public Sub SaveRTF()
filenum% = FreeFile
Open CurrentPath + "rtf.cfg" For Output As #filenum
Print #filenum, PText.Text
Print #filenum, ScrollText.Text
Close #filenum
RTFStartPosition = Val(PText.Text)
RTFStep = Val(ScrollText.Text)
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
     SaveRTF
     Unload RTFForm
Case 1
     Unload RTFForm
End Select
End Sub

Private Sub Form_Load()
Dim a As Variant
Dim b As Variant
filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "rtf.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Close #filenum
PText.Text = a
ScrollText.Text = b
RTFForm.Top = gtop + gheight
RTFForm.Left = 2500
Exit Sub
dealerror:
Close #filenum
PText.Text = "80"
ScrollText.Text = "10"
RTFStartPosition = 80
RTFStep = 10
RTFForm.Top = gtop + gheight
RTFForm.Left = 2500
End Sub

Private Sub PText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub

Private Sub ScrollText_KeyPress(KeyAscii As Integer)
TestKey KeyAscii
End Sub
