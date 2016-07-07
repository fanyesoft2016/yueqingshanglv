VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{288F1520-FAC4-11CE-B16F-00AA0060D93D}#1.0#0"; "MCIWNDX.OCX"
Begin VB.Form DisForm 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "ËÎÌו"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MCIWndX.MCIWnd MCIWnd1 
      Height          =   135
      Left            =   -600
      TabIndex        =   1
      Top             =   0
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   238
      _StockProps     =   96
      Menu            =   0   'False
      Playbar         =   0   'False
      Record          =   0   'False
      Repeat          =   0   'False
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   30
      Left            =   -100
      TabIndex        =   0
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393217
      BackColor       =   -2147483642
      BorderStyle     =   0
      Enabled         =   0   'False
      Appearance      =   0
      RightMargin     =   1
      TextRTF         =   $"tpmsplay3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌו"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   15
      Left            =   -300
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "DisForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

                    
Private Sub Form_Load()
 TopMost% = SetWindowPos(DisForm.hwnd, -1, 0, 0, 0, 0, wFlags)
End Sub

Private Sub MCIWnd1_PositionChange(ByVal Position As Long)
'CurrentPosition = Position
If Flag = True Then
  If Position = avilength Then
    If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
    Else
         FirstFlags = True
    End If
    DisForm.MCIWnd1.Command = "close"
    Call AppForm.SendNextMessage
   End If
End If
End Sub

