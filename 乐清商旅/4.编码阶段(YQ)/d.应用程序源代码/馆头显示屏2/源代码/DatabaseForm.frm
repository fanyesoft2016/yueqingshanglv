VERSION 5.00
Begin VB.Form DatabaseForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "12:23"
   ClientHeight    =   4485
   ClientLeft      =   495
   ClientTop       =   1755
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -2000
      Top             =   3210
   End
   Begin VB.Timer DisTimer 
      Enabled         =   0   'False
      Left            =   -2000
      Top             =   2130
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   11
      X1              =   7215
      X2              =   7215
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   6660
      X2              =   6660
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Label LblLasLeftSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   6825
      TabIndex        =   26
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "后"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   10
      Left            =   6840
      TabIndex        =   25
      Top             =   540
      Width           =   210
   End
   Begin VB.Label LblTomLeftSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   6315
      TabIndex        =   24
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   9
      Left            =   6315
      TabIndex        =   23
      Top             =   540
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "浙江方苑科技"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7170
      TabIndex        =   22
      Top             =   285
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "东阳西站车次售票情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   3120
      TabIndex        =   21
      Top             =   105
      Width           =   3765
   End
   Begin VB.Label TimerLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20:20:01"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1560
      TabIndex        =   20
      Top             =   285
      Width           =   720
   End
   Begin VB.Label DateLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1998年12月30日"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   285
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   165
      X2              =   8445
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   150
      X2              =   150
      Y1              =   495
      Y2              =   4305
   End
   Begin VB.Line LineEnd 
      BorderColor     =   &H000000FF&
      X1              =   150
      X2              =   8430
      Y1              =   4305
      Y2              =   4305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   10
      X1              =   8430
      X2              =   8430
      Y1              =   510
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   150
      X2              =   8445
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   840
      X2              =   840
      Y1              =   495
      Y2              =   4305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   1620
      X2              =   1620
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   7
      X1              =   4935
      X2              =   4935
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   6
      X1              =   4095
      X2              =   4095
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   4
      X1              =   2595
      X2              =   2595
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   8
      X1              =   5535
      X2              =   5535
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   9
      X1              =   6135
      X2              =   6135
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   270
      TabIndex        =   18
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   1
      Left            =   1020
      TabIndex        =   17
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终点站"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   2
      Left            =   1770
      TabIndex        =   16
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   3
      Left            =   3435
      TabIndex        =   15
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "全票价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   4
      Left            =   4245
      TabIndex        =   14
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   5
      Left            =   5025
      TabIndex        =   13
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "余票"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   6
      Left            =   5625
      TabIndex        =   12
      Top             =   540
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车属单位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   7
      Left            =   7425
      TabIndex        =   11
      Top             =   540
      Width           =   840
   End
   Begin VB.Label lblSBusID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12:23"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   945
      TabIndex        =   9
      Top             =   855
      Width           =   525
   End
   Begin VB.Label lblterminal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "百科全数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   1650
      TabIndex        =   8
      Top             =   855
      Width           =   930
   End
   Begin VB.Label Lalcar 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "普通车"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   3345
      TabIndex        =   7
      Top             =   855
      Width           =   630
   End
   Begin VB.Label Lbltprice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "139.2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   4275
      TabIndex        =   6
      Top             =   855
      Width           =   525
   End
   Begin VB.Label lblseatsum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   5115
      TabIndex        =   5
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Lblleftseat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   5715
      TabIndex        =   4
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Lblentrance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客运中心"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   7440
      TabIndex        =   3
      Top             =   855
      Width           =   840
   End
   Begin VB.Label lblWeekDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "星期一"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2385
      TabIndex        =   2
      Top             =   285
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   5
      X1              =   3195
      X2              =   3195
      Y1              =   510
      Y2              =   4290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "里程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   8
      Left            =   2685
      TabIndex        =   1
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lblMiter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   2715
      TabIndex        =   0
      Top             =   855
      Width           =   420
   End
End
Attribute VB_Name = "DatabaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim iRecordSum As Integer
Dim iRecordPoint As Integer
Dim nTmpRecordSet As Integer
Dim sCurrentRecordSet(23, 11) As String


Const cnDist = 240

Private Sub DisTimer_Timer()
GetDataRecord
DisplayDatabase
End Sub
Private Sub DisplayDatabase()
Dim i As Integer
Dim j As Integer
On Error Resume Next
Dim cl As Long
    For i = 0 To iScrollRowofRealTime - 1
    '颜色设定
        cl = vbRed
        lblSBusID(i).ForeColor = cl
        lblOffTime(i).ForeColor = cl
        lblterminal(i).ForeColor = cl
        Lalcar(i).ForeColor = cl
        Lbltprice(i).ForeColor = cl
        lblseatsum(i).ForeColor = cl
        Lblleftseat(i).ForeColor = cl
        Lblentrance(i).ForeColor = cl
        lblMiter(i).ForeColor = cl
        LblTomLeftSeat(i).ForeColor = cl
        LblLasLeftSeat(i).ForeColor = cl
        
        
        
    '车次日期
'        lblSBusDate(i).Caption = sCurrentRecordSet(i, 9)
    '车次
        lblSBusID(i).Caption = sCurrentRecordSet(i, 0)
    '发车时间
        lblOffTime(i).Caption = sCurrentRecordSet(i, 1)
    '终点站
        lblterminal(i) = sCurrentRecordSet(i, 2)
    '车型
        Lalcar(i) = sCurrentRecordSet(i, 3)
    '票价(终点站票价)
        Lbltprice(i) = sCurrentRecordSet(i, 4)
    '总座数
        lblseatsum(i) = sCurrentRecordSet(i, 5)
    '剩余票数
        Lblleftseat(i) = sCurrentRecordSet(i, 6)
    '车属单位
        Lblentrance(i) = sCurrentRecordSet(i, 7)
    '里程
        lblMiter(i) = sCurrentRecordSet(i, 8)
    '明天余座
        LblTomLeftSeat(i) = sCurrentRecordSet(i, 10)
    '后天余座
        LblLasLeftSeat(i) = sCurrentRecordSet(i, 11)
    Next

End Sub

Private Sub Form_Load()
    Dim TopMost As Integer
    Dim szSql As String
    Dim i As Integer
    Dim j As Integer
    
    Dim clTemp As Long
    
    '设置界面的线的长度及位置
    clTemp = lblSBusID(0).Top + cnDist * iScrollRowofRealTime + 8
    For i = 0 To 11
        Line2(i).Y2 = clTemp
        
    Next i
    LineEnd.y1 = clTemp
    LineEnd.Y2 = clTemp
    'iScrollRowofRealTime
    '动态装载控件
    For i = 1 To iScrollRowofRealTime - 1
        Load lblSBusID(i)
        lblSBusID(i).Top = lblSBusID(0).Top + cnDist * i
        lblSBusID(i).Left = lblSBusID(0).Left
        lblSBusID(i).Visible = True
        
        Load lblOffTime(i)
        lblOffTime(i).Top = lblOffTime(0).Top + cnDist * i
        lblOffTime(i).Left = lblOffTime(0).Left
        lblOffTime(i).Visible = True
        
        Load lblterminal(i)
        lblterminal(i).Top = lblterminal(0).Top + cnDist * i
        lblterminal(i).Left = lblterminal(0).Left
        lblterminal(i).Visible = True
        lblterminal(i).Alignment = lblterminal(0).Alignment
        
        
        Load Lalcar(i)
        Lalcar(i).Top = Lalcar(0).Top + cnDist * i
        Lalcar(i).Left = Lalcar(0).Left
        Lalcar(i).Visible = True
        
        Load Lbltprice(i)
        Lbltprice(i).Top = Lbltprice(0).Top + cnDist * i
        Lbltprice(i).Left = Lbltprice(0).Left
        Lbltprice(i).Visible = True
        
        Load lblseatsum(i)
        lblseatsum(i).Top = lblseatsum(0).Top + cnDist * i
        lblseatsum(i).Left = lblseatsum(0).Left
        lblseatsum(i).Visible = True
        
        Load Lblleftseat(i)
        Lblleftseat(i).Top = Lblleftseat(0).Top + cnDist * i
        Lblleftseat(i).Left = Lblleftseat(0).Left
        Lblleftseat(i).Visible = True
        
        Load Lblentrance(i)
        Lblentrance(i).Top = Lblentrance(0).Top + cnDist * i
        Lblentrance(i).Left = Lblentrance(0).Left
        Lblentrance(i).Visible = True
        
        Load lblMiter(i)
        lblMiter(i).Top = lblMiter(0).Top + cnDist * i
        lblMiter(i).Left = lblMiter(0).Left
        lblMiter(i).Visible = True
        
        Load LblTomLeftSeat(i)
        LblTomLeftSeat(i).Top = LblTomLeftSeat(0).Top + cnDist * i
        LblTomLeftSeat(i).Left = LblTomLeftSeat(0).Left
        LblTomLeftSeat(i).Visible = True
        
        
        Load LblLasLeftSeat(i)
        LblLasLeftSeat(i).Top = LblLasLeftSeat(0).Top + cnDist * i
        LblLasLeftSeat(i).Left = LblLasLeftSeat(0).Left
        LblLasLeftSeat(i).Visible = True
        
        
    Next i
    
    Const wFlags1 = &H2 Or &H1 Or &H40 Or &H10

    DisplayDatabase
    Timer1.Interval = 1000
    DisTimer.Interval = iIntervaltimeofRealTime * 1000
    Timer1.Enabled = True
    
    
    lblTitle.Caption = frmSetTitle.txtTitle.Text
    lblTitle.Left = Val(frmSetTitle.XText.Text)
    lblTitle.Top = Val(frmSetTitle.YText.Text)
    lblWeekDay.Caption = WeekdayName(Weekday(Date))
    
    TopMost = SetWindowPos(DatabaseForm.hwnd, -1, 0, 0, 0, 0, wFlags1)
    
    On Error GoTo here
    Set rst = GetDisplay(oSystem.NowDateTime)
    iRecordPoint = 0
    iRecordSum = rst.RecordCount
    rst.MoveFirst
    GetDataRecord
    DisplayDatabase
    DisTimer.Enabled = True
    Exit Sub
here:
    DisTimer.Enabled = False
    Timer1.Enabled = False
    Dim action As Integer
    action = MsgBox("读取服务器出错或未取得数据，不能进行实时车次售票状况显示,请查询数据库再试!" & vbCrLf & "原始信息:" & ErrDescription, vbInformation, "出错")
    flagofdatabaseform = False
Unload DatabaseForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
DisTimer.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim da As Date
    da = oSystem.NowDateTime
    TimerLabel.Caption = Format(da, "HH:MM:SS")
    DateLabel.Caption = Format(da, "YYYY年MM月DD日")
    lblWeekDay.Caption = WeekdayName(Weekday(da))
End Sub

Private Sub GetDataRecord()
Dim i, j As Integer
Dim iCurrentIndex As Integer
iCurrentIndex = iRecordPoint
'If iCurrentIndex <= rst.RecordCount Then
'
'Else
'    GetData
'End If
'On Error Resume Next

For i = 0 To iScrollRowofRealTime
    For j = 0 To 11
        sCurrentRecordSet(i, j) = ""
    Next j
Next i
If rst.RecordCount <= iScrollRowofRealTime Then
    GetData
    For i = 0 To rst.RecordCount - 1
        If rst.EOF Then
            GetData
        End If
        
'szSql = "SELECT EBW.bus_id " & "," _
'        & "EBW.check_gate_id " & "," _
'        & "EBW.vehicle_type_name " & "," _
'        & "EBW.total_seat " & "," _
'        & "EBW.sale_seat_quantity " & "," _
'        & "EBW.bus_start_time " & "," _
'        & "EBW.end_station_name " & "," _
'        & "TCI.transport_company_short_name " & "," _
'        & "EPW.ticket_price_total " & "," _
'        & "SI.station_name," _
'        & "ri.mileage "
        
        
        
        
        Dim szStationName As String
        
        sCurrentRecordSet(i, 0) = rst.Fields("bus_id").Value
        szStationName = rst.Fields("station_name").Value
        If LenA(szStationName) = 5 Or LenA(szStationName) = 6 Then
            sCurrentRecordSet(i, 2) = " " & rst.Fields("station_name").Value & " "
        ElseIf LenA(szStationName) = 3 Or LenA(szStationName) = 4 Then
            sCurrentRecordSet(i, 2) = "  " & rst.Fields("station_name").Value & "  "
        ElseIf LenA(szStationName) = 7 Or LenA(szStationName) = 8 Then
            sCurrentRecordSet(i, 2) = rst.Fields("station_name").Value
        End If
        sCurrentRecordSet(i, 3) = Left(rst.Fields("vehicle_type_name").Value, 3)
        sCurrentRecordSet(i, 4) = rst.Fields("ticket_price_total").Value
        If sCurrentRecordSet(i, 4) > 100 Then
            sCurrentRecordSet(i, 4) = Format(sCurrentRecordSet(i, 4), "0.0")
        Else
            sCurrentRecordSet(i, 4) = Format(sCurrentRecordSet(i, 4), "0.00")
        End If
        
        
        sCurrentRecordSet(i, 7) = rst.Fields("transport_company_name").Value
        sCurrentRecordSet(i, 8) = rst.Fields("mileage").Value
        sCurrentRecordSet(i, 9) = Format(rst.Fields("bus_date").Value, "mm-dd")
        
'        .Append "tom_sale_seat_quantity", adSmallInt
'        .Append "las_sale_seat_quantity", adSmallInt
        
        '若为流水车次
        If rst.Fields("bus_type").Value = TP_ScrollBus Then
            sCurrentRecordSet(i, 1) = "滚动"
            sCurrentRecordSet(i, 5) = "滚动"
            sCurrentRecordSet(i, 6) = ""
            sCurrentRecordSet(i, 10) = ""
            sCurrentRecordSet(i, 11) = ""
        Else
            sCurrentRecordSet(i, 1) = Format(rst.Fields("bus_start_time").Value, "HH:MM")
            sCurrentRecordSet(i, 5) = rst.Fields("total_seat").Value
            sCurrentRecordSet(i, 6) = rst.Fields("sale_seat_quantity").Value
            sCurrentRecordSet(i, 10) = rst.Fields("tom_sale_seat_quantity").Value
            sCurrentRecordSet(i, 11) = rst.Fields("las_sale_seat_quantity").Value
        End If
        rst.MoveNext
        iCurrentIndex = iCurrentIndex + 1
    Next
    iRecordPoint = iRecordPoint + iScrollRowofRealTime
    iRecordPoint = iRecordPoint Mod iRecordSum
    Exit Sub
    
Else
    For i = 0 To iScrollRowofRealTime - 1
        '   If iCurrentIndex >= iRecordSum Then
        If rst.EOF Then
            GetData
        End If
        '   End If
        sCurrentRecordSet(i, 0) = rst.Fields("bus_id").Value
        'sCurrentRecordSet(i, 2) = rst.Fields("station_name").Value
        
        szStationName = Trim(rst.Fields("station_name").Value)
        If LenA(szStationName) = 5 Or LenA(szStationName) = 6 Then
            sCurrentRecordSet(i, 2) = " " & rst.Fields("station_name").Value & " "
        ElseIf LenA(szStationName) = 3 Or LenA(szStationName) = 4 Then
            sCurrentRecordSet(i, 2) = "  " & rst.Fields("station_name").Value & "  "
        ElseIf LenA(szStationName) = 7 Or LenA(szStationName) = 8 Then
            sCurrentRecordSet(i, 2) = rst.Fields("station_name").Value
        Else
            sCurrentRecordSet(i, 2) = rst.Fields("station_name").Value
        End If
        sCurrentRecordSet(i, 3) = Left(rst.Fields("vehicle_type_name").Value, 3)
        sCurrentRecordSet(i, 4) = rst.Fields("ticket_price_total").Value
        If sCurrentRecordSet(i, 4) > 100 Then
            sCurrentRecordSet(i, 4) = Format(sCurrentRecordSet(i, 4), "000.00")
        Else
            sCurrentRecordSet(i, 4) = Format(sCurrentRecordSet(i, 4), "0.00")
        End If
        sCurrentRecordSet(i, 7) = rst.Fields("transport_company_name").Value
        sCurrentRecordSet(i, 8) = rst.Fields("mileage").Value
'        sCurrentRecordSet(i, 9) = Format(rst.Fields("bus_date").Value, "mm-dd")
        
        '若为流水车次
        If rst.Fields("bus_type").Value = "True" Then
            sCurrentRecordSet(i, 1) = "滚动"
            sCurrentRecordSet(i, 5) = "滚动"
            sCurrentRecordSet(i, 6) = ""
            sCurrentRecordSet(i, 10) = ""
            sCurrentRecordSet(i, 11) = ""
        Else
            sCurrentRecordSet(i, 1) = Format(rst.Fields("bus_start_time").Value, "HH:MM")
            sCurrentRecordSet(i, 5) = rst.Fields("total_seat").Value
            sCurrentRecordSet(i, 6) = rst.Fields("sale_seat_quantity").Value
            sCurrentRecordSet(i, 10) = rst.Fields("tom_sale_seat_quantity").Value
            sCurrentRecordSet(i, 11) = rst.Fields("las_sale_seat_quantity").Value
        End If
        rst.MoveNext
        iCurrentIndex = iCurrentIndex + 1
    Next
    iRecordPoint = iRecordPoint + iScrollRowofRealTime
    iRecordPoint = iRecordPoint Mod iRecordSum
    Exit Sub
End If
End Sub

Private Sub GetData()
    Dim iCurrentIndex As Integer
    On Error GoTo here
    Dim szNow As String
    szNow = ToDBDateTime(oSystem.NowDateTime)
    Set rst = GetDisplay(szNow) '"2000-04-05 " & Format(Time, "hh:mm:dd"))

    
    
    rst.MoveFirst
    iRecordSum = rst.RecordCount
    iCurrentIndex = 0
    Exit Sub
here:
    DisTimer.Enabled = False
    Timer1.Enabled = False
    Dim action As Integer
    action = MsgBox("读取服务器出错或未取得数据，不能进行实时车次售票状况显示,请查询数据库再试!" & vbCrLf & "原始信息:" & ErrDescription, vbInformation, "出错")
    flagofdatabaseform = False
    Unload DatabaseForm
    End
End Sub

Public Function ErrDescription()
    ErrDescription = err.Number & "->" & err.Description
End Function

Private Function GetDisplay(DayToSell As String) As Recordset
'    Static rsBusBook As Variant
'    Dim rsTempBook As Variant
    Static First As Boolean
    Dim rsDisplay As New Recordset
    Dim rsTemp As Recordset
    Dim szSql As String
    
    On Error GoTo err
    
'    szSql = " EXEC GetBigDisplay '" & ToDBDateTime(DayToSell) & "'"
        szSql = " SELECT ebw.bus_date , EBW.bus_id , c.check_gate_name , EBW.vehicle_type_name , EBW.total_seat ," _
& "    EBW.sale_seat_quantity ," _
& "    EBW.bus_start_time ," _
& "    EBW.end_station_name ," _
& "    TCI.transport_company_short_name transport_company_name," _
& "    EPW.ticket_price_total ," _
& "    ebw.bus_type ," _
& "    ri.mileage" _
& "    FROM work_env_bus_info EBW ,work_env_bus_price_lst EPW,company_info TCI,station_info AS SI" _
& "    , checkgate_info c  ,area_code ac,route_info ri" _
& "    WHERE DATEDIFF(day,convert(datetime,'" & DayToSell & "',120),EBW.bus_date)<3 " _
& "        AND DATEDIFF(day,convert(datetime,'" & DayToSell & "',120),EBW.bus_date)>=0" _
& "         AND DATEDIFF(day,EPW.bus_date,EBW.bus_date)=0" _
& "            AND EBW.bus_id = EPW.bus_id" _
& "            AND TCI.transport_company_id = EBW.transport_company_id" _
& "        AND EBW.end_station_id=EPW.station_id" _
& "            AND EPW.ticket_type=1 and si.station_id = ebw.end_station_id" _
& "            AND ebw.check_gate_id = c.check_gate_id AND EPW.sell_station_id = c.sell_station_id" _
& "        and c.sell_station_id =  'xz'" _
& "            AND EBW.status <> 1" _
& "            AND EBW.route_id=ri.route_id" _
& "            AND ac.area_code=si.area_code" _
& "            ORDER BY convert(char(5), EBW.bus_start_time ,108), EBW.bus_id,ebw.bus_date  "


    Set rsTemp = oTo.Execute(szSql)
    ' 如果记录数为0,则终止
    If rsTemp.RecordCount = 0 Then End



    '将返回的记录集,同一车次3天的记录,转换到一条记录上去,将明天\后天的余座,作为字段列出来
    
    
    
    Dim i As Integer
    
    Dim szBusID As String
    Dim szCheckGateName As String
    Dim szVehicleTypeName As String
    Dim nTotalSeat As Integer
    Dim nSaleSeatQuantity As Integer
    Dim nTomSaleSeatQuantity As Integer
    Dim nLasSaleSeatQuantity  As Integer
    Dim dtBusStartTime As Date
    Dim szCompanyName As String
    Dim dbTicketPrice As Double
    Dim szStationName As String
    Dim nBusType As Integer
    Dim dbMileage As Double
    
    
    
    'bus_date                                               bus_id check_gate_name                                    vehicle_type_name total_seat
    'sale_seat_quantity bus_start_time                                         end_station_name transport_company_name
    'ticket_price_total station_name bus_type mileage
    With rsDisplay.Fields
        
        .Append "bus_id", adChar, 5
        .Append "check_gate_name", adVarChar, 50
        .Append "vehicle_type_name", adVarChar, 10
        .Append "total_seat", adSmallInt
        
        
        .Append "sale_seat_quantity", adSmallInt
        .Append "tom_sale_seat_quantity", adSmallInt
        .Append "las_sale_seat_quantity", adSmallInt
        .Append "bus_start_time", adDBTime
        .Append "transport_company_name", adChar, 10
        .Append "ticket_price_total", adDouble
        .Append "station_name", adChar, 10
        .Append "bus_type", adSmallInt
        .Append "mileage", adDouble
        
    End With
    
    rsTemp.MoveFirst
    '赋初值
    szBusID = FormatDbValue(rsTemp!bus_id)
    szCheckGateName = FormatDbValue(rsTemp!check_gate_name)
    szVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)
    nTotalSeat = FormatDbValue(rsTemp!total_seat)
    dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
    '判断日期为什么时候,如果为当天,则放到当天的余票中,明天则放明天中
    
    Dim nDateDiff As Integer
    nDateDiff = DateDiff("d", Date, FormatDbValue(rsTemp!bus_date))
    If nDateDiff = 0 Then
        '如果发车时间已过,则余座要变为0
        If Format(Now, "hh:mm") > Format(FormatDbValue(rsTemp!bus_start_time), "hh:mm") Then
            
            nSaleSeatQuantity = 0
        Else
            nSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
        End If
    ElseIf nDateDiff = 1 Then
        nTomSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
    ElseIf nDateDiff = 2 Then
        nLasSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
    End If
    
    dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
    szCompanyName = FormatDbValue(rsTemp!transport_company_name)
    dbTicketPrice = FormatDbValue(rsTemp!ticket_price_total)
    szStationName = FormatDbValue(rsTemp!end_station_name)
    nBusType = FormatDbValue(rsTemp!bus_type)
    dbMileage = FormatDbValue(rsTemp!mileage)
    
    
    
    rsDisplay.Open
    For i = 1 To rsTemp.RecordCount
        If szBusID <> FormatDbValue(rsTemp!bus_id) Then
            '新增一条记录
            rsDisplay.AddNew
            rsDisplay!bus_id = szBusID
            rsDisplay!check_gate_name = szCheckGateName
            rsDisplay!vehicle_type_name = szVehicleTypeName
            rsDisplay!total_seat = nTotalSeat
            rsDisplay!bus_start_time = dtBusStartTime
            rsDisplay!sale_seat_quantity = nSaleSeatQuantity
            rsDisplay!tom_sale_seat_quantity = nTomSaleSeatQuantity
            rsDisplay!las_sale_seat_quantity = nLasSaleSeatQuantity
            rsDisplay!transport_company_name = szCompanyName
            rsDisplay!ticket_price_total = dbTicketPrice
            rsDisplay!station_name = szStationName
            rsDisplay!bus_type = nBusType
            rsDisplay!mileage = dbMileage
            rsDisplay.Update
            nTomSaleSeatQuantity = 0
            nLasSaleSeatQuantity = 0
            '赋初值
            szBusID = FormatDbValue(rsTemp!bus_id)
            szCheckGateName = FormatDbValue(rsTemp!check_gate_name)
            szVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)
            nTotalSeat = FormatDbValue(rsTemp!total_seat)
            dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
            '判断日期为什么时候,如果为当天,则放到当天的余票中,明天则放明天中
            
            nDateDiff = DateDiff("d", Date, FormatDbValue(rsTemp!bus_date))
            If nDateDiff = 0 Then
                '如果发车时间已过,则余座要变为0
                If Format(Now, "hh:mm") > Format(FormatDbValue(rsTemp!bus_start_time), "hh:mm") Then
                    
                    nSaleSeatQuantity = 0
                Else
                    nSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
                End If
            ElseIf nDateDiff = 1 Then
                nTomSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
            ElseIf nDateDiff = 2 Then
                nLasSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
            End If
            
            dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
            szCompanyName = FormatDbValue(rsTemp!transport_company_name)
            dbTicketPrice = FormatDbValue(rsTemp!ticket_price_total)
            szStationName = FormatDbValue(rsTemp!end_station_name)
            nBusType = FormatDbValue(rsTemp!bus_type)
            dbMileage = FormatDbValue(rsTemp!mileage)
            
            
        Else
            
            nDateDiff = DateDiff("d", Date, FormatDbValue(rsTemp!bus_date))
            If nDateDiff = 0 Then
                '如果发车时间已过,则余座要变为0
                If Format(Now, "hh:mm") > Format(FormatDbValue(rsTemp!bus_start_time), "hh:mm") Then
                    
                    nSaleSeatQuantity = 0
                Else
                    nSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
                End If
            ElseIf nDateDiff = 1 Then
                nTomSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
            ElseIf nDateDiff = 2 Then
                nLasSaleSeatQuantity = FormatDbValue(rsTemp!sale_seat_quantity)
            End If
        End If
        rsTemp.MoveNext
            
    Next i

    rsDisplay.AddNew
    rsDisplay!bus_id = szBusID
    rsDisplay!check_gate_name = szCheckGateName
    rsDisplay!vehicle_type_name = szVehicleTypeName
    rsDisplay!total_seat = nTotalSeat
    rsDisplay!bus_start_time = dtBusStartTime
    rsDisplay!sale_seat_quantity = nSaleSeatQuantity
    rsDisplay!tom_sale_seat_quantity = nTomSaleSeatQuantity
    rsDisplay!las_sale_seat_quantity = nLasSaleSeatQuantity
    rsDisplay!transport_company_name = szCompanyName
    rsDisplay!ticket_price_total = dbTicketPrice
    rsDisplay!station_name = szStationName
    rsDisplay!bus_type = nBusType
    rsDisplay!mileage = dbMileage
    rsDisplay.Update
            
    
    
    

    
    Set GetDisplay = rsDisplay
Exit Function
err:
    MsgBox err.Description, vbInformation, err.Number
End Function



'是否显示
Private Sub RecordSetIsDisp(Optional IsDisp As Boolean = False, Optional nRecordCount As Integer = 10)
Dim i As Integer
For i = 0 To nRecordCount
    lblSBusID(i).Visible = IsDisp
    lblOffTime(i).Visible = IsDisp
    lblterminal(i).Visible = IsDisp
    Lalcar(i).Visible = IsDisp
    Lbltprice(i).Visible = IsDisp
    lblseatsum(i).Visible = IsDisp
    Lblleftseat(i).Visible = IsDisp
    Lblentrance(i).Visible = IsDisp
    lblMiter(i).Visible = IsDisp
    LblTomLeftSeat(i).Visible = IsDisp
    LblLasLeftSeat(i).Visible = IsDisp
'    lblSBusDate(i).Visible = IsDisp
    
Next i
End Sub

'**************************************



