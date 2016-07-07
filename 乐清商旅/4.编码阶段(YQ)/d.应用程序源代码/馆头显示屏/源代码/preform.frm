VERSION 5.00
Begin VB.Form PreView 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Picture         =   "preform.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4215
      Top             =   1995
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   5220
      X2              =   120
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   135
      X2              =   135
      Y1              =   120
      Y2              =   2985
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   135
      X2              =   5205
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   5220
      X2              =   5220
      Y1              =   120
      Y2              =   2985
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Height          =   2880
      Left            =   135
      Top             =   120
      Width           =   5100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "播 放 软 件 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   1635
      TabIndex        =   1
      Top             =   1170
      Width           =   2940
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "L E D 多 媒 体 屏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   1545
      TabIndex        =   0
      Top             =   420
      Width           =   3540
   End
End
Attribute VB_Name = "PreView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Load AppForm
End Sub


Private Sub Label7_Click()

End Sub

Private Sub Timer1_Timer()
Unload PreView
AppForm.Show
End Sub
