VERSION 5.00
Begin VB.Form DatabaseForm 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "12:23"
   ClientHeight    =   9120
   ClientLeft      =   75
   ClientTop       =   3840
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11725.72
   ScaleMode       =   0  'User
   ScaleWidth      =   16320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6330
      Top             =   5610
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   4590
   End
   Begin VB.Timer DisTimer 
      Enabled         =   0   'False
      Left            =   -2025
      Top             =   2130
   End
   Begin VB.Label TimerLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09:20:01"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label DateLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2006��07��30��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   2430
   End
   Begin VB.Label lblWeekDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5040
      TabIndex        =   14
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   13890
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblStation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���ݻ�վ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Index           =   0
      Left            =   9120
      TabIndex        =   12
      Top             =   1920
      Width           =   4980
   End
   Begin VB.Label lblOffTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18:00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   1920
      Width           =   3420
   End
   Begin VB.Label lblBusID 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "88888��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Index           =   0
      Left            =   195
      TabIndex        =   10
      Top             =   1920
      Width           =   3420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�յ�վ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Index           =   2
      Left            =   10440
      TabIndex        =   9
      Top             =   720
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Top             =   720
      Width           =   3840
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Index           =   0
      Left            =   825
      TabIndex        =   7
      Top             =   690
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʊ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   8
      Left            =   18615
      TabIndex        =   6
      Top             =   5010
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblCheckGate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   18615
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "״̬"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Index           =   4
      Left            =   15480
      TabIndex        =   4
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ";��վ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Index           =   5
      Left            =   15480
      TabIndex        =   3
      Top             =   240
      Width           =   2880
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���ڼ�Ʊ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Index           =   0
      Left            =   15480
      TabIndex        =   2
      Top             =   3840
      Width           =   3420
   End
   Begin VB.Label lblPassStation 
      BackStyle       =   0  'Transparent
      Caption         =   "�鰲,��ũ��,С�Ŵ�,�縣,����,��ͬ,����,��ƽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Index           =   0
      Left            =   15450
      TabIndex        =   1
      Top             =   1320
      Width           =   25350
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   6
      X1              =   22515
      X2              =   22515
      Y1              =   -289.286
      Y2              =   4532.145
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   5
      X1              =   15300
      X2              =   15300
      Y1              =   790.715
      Y2              =   5612.146
   End
   Begin VB.Line LineEnd 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   60
      X2              =   15300
      Y1              =   5612.146
      Y2              =   5612.146
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   4
      X1              =   15840
      X2              =   15840
      Y1              =   790.715
      Y2              =   5612.146
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   60
      X2              =   60
      Y1              =   790.715
      Y2              =   5631.431
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ͷ�����������ļ�Ʊ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   6840
      TabIndex        =   0
      Top             =   0
      Width           =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   90
      X2              =   15300
      Y1              =   790.715
      Y2              =   790.715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   60
      X2              =   15300
      Y1              =   2372.144
      Y2              =   2372.144
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   3870
      X2              =   3870
      Y1              =   790.715
      Y2              =   5631.431
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   8550
      X2              =   8550
      Y1              =   790.715
      Y2              =   5612.146
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   15555
      X2              =   15555
      Y1              =   3259.287
      Y2              =   8080.718
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
Dim sCurrentRecordSet(40, 6) As String

Dim aszScrollBus() As Variant
Dim nScrollIndex As Integer
Dim nScrollCount As Integer

Dim bScroll As Boolean

Const cnDist = 1200 '465
Const cnSist = 337

Const cnDist2 = 11550 '��������ĺ������

Dim nCount As Integer
Dim nCount2 As Integer

'**************************************
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
    '��ɫ�趨
        If sCurrentRecordSet(i, 4) = "���ڼ�Ʊ" Or sCurrentRecordSet(i, 4) = "���ڲ���" Or Val(sCurrentRecordSet(i, 6)) = 1 Then
            cl = vbRed
            lblStatus(i).Tag = 1
'            lblBusID(i).Tag = 1
'            lblOffTime(i).Tag = 1
'            lblStation(i).Tag = 1
'            lblCheckGate(i).Tag = 1
'            lblPassStation(i).Tag = 1
        ElseIf sCurrentRecordSet(i, 4) = "ֹͣ��Ʊ" Then
            cl = vbRed
            lblStatus(i).Tag = 1
'            lblBusID(i).Tag = 1
'            lblOffTime(i).Tag = 1
'            lblStation(i).Tag = 1
'            lblCheckGate(i).Tag = 1
'            lblPassStation(i).Tag = 1
        Else
            cl = vbRed
            lblStatus(i).Tag = 0
'            lblBusID(i).Tag = 0
'            lblOffTime(i).Tag = 0
'            lblStation(i).Tag = 0
'            lblCheckGate(i).Tag = 0
'            lblPassStation(i).Tag = 0
            
            lblStatus(i).Visible = True
'            lblBusID(i).Visible = True
'            lblOffTime(i).Visible = True
'            lblStation(i).Visible = True
'            lblCheckGate(i).Visible = True
'            lblPassStation(i).Visible = True
        End If
        
        lblBusID(i).ForeColor = cl
        lblOffTime(i).ForeColor = cl
        lblStation(i).ForeColor = cl
        lblCheckGate(i).ForeColor = cl
        lblStatus(i).ForeColor = cl
        lblPassStation(i).ForeColor = cl

    '����
        lblBusID(i).Caption = Trim(sCurrentRecordSet(i, 0))
    '����ʱ��
        lblOffTime(i).Caption = sCurrentRecordSet(i, 1)
    '�յ�վ
        lblStation(i) = sCurrentRecordSet(i, 2)
    '��Ʊ��
        lblCheckGate(i) = sCurrentRecordSet(i, 3)
    '״̬
        lblStatus(i) = sCurrentRecordSet(i, 4)
    ';��վ
        lblPassStation(i) = sCurrentRecordSet(i, 5)
    Next

End Sub

Private Sub Form_Load()
    Dim TopMost As Integer
    Dim szSql As String
    Dim i As Integer
    Dim j As Integer
    
    Dim clTemp As Long
    Dim lLeftDist As Long '����ĺ������,��ҪΪ���������ʾ��
    Dim lTopDist As Long
    
    '���ý�����ߵĳ��ȼ�λ��
    clTemp = lblBusID(0).Top + cnDist * iScrollRowofRealTime + 50
    For i = 0 To 6
        Line2(i).Y2 = clTemp
    Next i
    

    LineEnd(0).y1 = clTemp
    LineEnd(0).Y2 = clTemp
    
'    LineEnd(1000).y1 = clTemp
'    LineEnd(1000).Y2 = clTemp
    
    'iScrollRowofRealTime
    '��̬װ�ؿؼ�
    For i = 1 To iScrollRowofRealTime - 1
'        lLeftDist = (i Mod 2) * cnDist2 '���Ϊż���������Ӧ����
'        If i >= iScrollRowofRealTime / 2 Then
'            lLeftDist = cnDist2
'            lTopDist = cnDist * (i - iScrollRowofRealTime / 2)
'        Else
            lLeftDist = 0
            lTopDist = cnDist * i
'        End If
        
        Load lblBusID(i)
        lblBusID(i).Top = lblBusID(0).Top + lTopDist
        lblBusID(i).Left = lblBusID(0).Left + lLeftDist
        lblBusID(i).Visible = True
        
        Load LineEnd(i)
        LineEnd(i).x1 = LineEnd(0).x1 + lLeftDist
        LineEnd(i).X2 = LineEnd(0).X2 + lLeftDist
        LineEnd(i).y1 = lblBusID(i).Top - 50
        LineEnd(i).Y2 = LineEnd(i).y1
        LineEnd(i).Visible = True
        
        Load lblOffTime(i)
        lblOffTime(i).Top = lblOffTime(0).Top + lTopDist
        lblOffTime(i).Left = lblOffTime(0).Left + lLeftDist
        lblOffTime(i).Visible = True
        
        Load lblStation(i)
        lblStation(i).Top = lblStation(0).Top + lTopDist
        lblStation(i).Left = lblStation(0).Left + lLeftDist
        lblStation(i).Visible = True
        
        
        Load lblCheckGate(i)
        lblCheckGate(i).Top = lblCheckGate(0).Top + lTopDist
        lblCheckGate(i).Left = lblCheckGate(0).Left + lLeftDist
        lblCheckGate(i).Visible = True
        
        Load lblStatus(i)
        lblStatus(i).Top = lblStatus(0).Top + lTopDist
        lblStatus(i).Left = lblStatus(0).Left + lLeftDist
        lblStatus(i).Visible = True
        
        Load lblPassStation(i)
        lblPassStation(i).Top = lblPassStation(0).Top + lTopDist
        lblPassStation(i).Left = lblPassStation(0).Left + lLeftDist
        lblPassStation(i).Visible = True
    Next i


    Const wFlags1 = &H2 Or &H1 Or &H40 Or &H10

'    DisplayDatabase
    Timer1.Interval = 1000
    DisTimer.Interval = iIntervaltimeofRealTime * 1000
    Timer1.Enabled = True
    
    
    lblTitle.Caption = frmSetTitle.txtTitle.Text
    'lblTitle.Left = Val(frmSetTitle.XText.Text)
    lblTitle.Top = Val(frmSetTitle.YText.Text)
    lblWeekDay.Caption = WeekdayName(Weekday(Date))
    
    TopMost = SetWindowPos(DatabaseForm.hwnd, -1, 0, 0, 0, 0, wFlags1)
    
    On Error GoTo here
    Set rst = GetDisplay(oSystem.NowDateTime)
    iRecordPoint = 0
    iRecordSum = rst.RecordCount
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        GetDataRecord
        DisplayDatabase
    End If
    
    DisTimer.Enabled = True
    Exit Sub
here:
    DisTimer.Enabled = False
    Timer1.Enabled = False
    Dim action As Integer
    action = MsgBox("��ȡ�����������δȡ�����ݣ����ܽ���ʵʱ������Ʊ״����ʾ,���ѯ���ݿ�����!" & vbCrLf & "ԭʼ��Ϣ:" & ErrDescription, vbInformation, "����")
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
    DateLabel.Caption = Format(da, "YYYY��MM��DD��")
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
    For j = 0 To 6
        sCurrentRecordSet(i, j) = ""
    Next j
Next i
If rst.RecordCount <= iScrollRowofRealTime Then
    GetData
    For i = 0 To rst.RecordCount - 1
        If rst.EOF Then
            GetData
        End If
        
        sCurrentRecordSet(i, 0) = rst.Fields("bus_id").Value
        sCurrentRecordSet(i, 2) = Trim(rst.Fields("end_station_name").Value)
        sCurrentRecordSet(i, 3) = Trim(rst.Fields("check_gate_id").Value)
        sCurrentRecordSet(i, 4) = ToBusStatusString(rst.Fields("status").Value)
        sCurrentRecordSet(i, 5) = rst.Fields("pass_station").Value
       ' sCurrentRecordSet(i, 6) = rst.Fields("hand_work").Value
        
        '��Ϊ��ˮ����
        If rst.Fields("bus_kind").Value = TP_ScrollBus Then
            sCurrentRecordSet(i, 1) = "����"
        Else
            sCurrentRecordSet(i, 1) = Format(rst.Fields("bus_start_time").Value, "HH:MM")
        End If
        rst.MoveNext
        iCurrentIndex = iCurrentIndex + 1
    Next
    iRecordPoint = iRecordPoint + iScrollRowofRealTime
    If iRecordSum > 0 Then
        iRecordPoint = iRecordPoint Mod iRecordSum
    End If
    Exit Sub
    
Else
    For i = 0 To iScrollRowofRealTime - 1
        '   If iCurrentIndex >= iRecordSum Then
        If rst.EOF Then
            GetData
        End If
        
        sCurrentRecordSet(i, 2) = Trim(rst.Fields("end_station_name").Value)
        sCurrentRecordSet(i, 3) = Trim(rst.Fields("check_gate_id").Value)
        sCurrentRecordSet(i, 4) = ToBusStatusString(rst.Fields("status").Value)
        sCurrentRecordSet(i, 5) = rst.Fields("pass_station").Value
        'sCurrentRecordSet(i, 6) = rst.Fields("hand_work").Value
        
        '��Ϊ��ˮ����
        If rst.Fields("bus_kind").Value = TP_ScrollBus Then  '"True" Then
            sCurrentRecordSet(i, 0) = rst.Fields("bus_id").Value
            sCurrentRecordSet(i, 1) = "����"
        Else
            sCurrentRecordSet(i, 0) = rst.Fields("bus_id").Value
            sCurrentRecordSet(i, 1) = Format(rst.Fields("bus_start_time").Value, "HH:MM")
        End If
        rst.MoveNext
        iCurrentIndex = iCurrentIndex + 1
    Next
    iRecordPoint = iRecordPoint + iScrollRowofRealTime
    If iRecordSum > 0 Then
        iRecordPoint = iRecordPoint Mod iRecordSum
    End If
    Exit Sub
End If
End Sub

Private Sub GetData()
    Dim iCurrentIndex As Integer
    On Error GoTo here
    Dim szNow As String
    szNow = ToDBDateTime(oSystem.NowDateTime)
    Set rst = GetDisplay(szNow) '"2000-04-05 " & Format(Time, "hh:mm:dd"))
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
    End If
    iRecordSum = rst.RecordCount
    iCurrentIndex = 0
    Exit Sub
here:
    Resume Next
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
    
    oTo.BeginTrans
    
'    oTo.Execute " EXEC GetXiaoPing "
'
   szSql = " EXEC cmGetDisplayData "

    Set rsTemp = oTo.Execute(szSql)
    
    oTo.CommitTrans
    ' �����¼��Ϊ0,����ֹ
'    If rsTemp.RecordCount = 0 Then End

    Set GetDisplay = rsTemp
Exit Function
err:
    MsgBox err.Description, vbInformation, err.Number
End Function



'�Ƿ���ʾ
Private Sub RecordSetIsDisp(Optional IsDisp As Boolean = False, Optional nRecordCount As Integer = 10)
Dim i As Integer
For i = 0 To nRecordCount
    lblBusID(i).Visible = IsDisp
    lblOffTime(i).Visible = IsDisp
    lblStation(i).Visible = IsDisp
    lblCheckGate(i).Visible = IsDisp
    lblStatus(i).Visible = IsDisp
    lblPassStation(i).Visible = IsDisp
Next i
End Sub

Private Sub FindRTF()
    Dim szDir As String
    nCount = 1
    Dim i As Integer
    Do While True
        
        szDir = CurrentPath + "����" & nCount & ".rtf"
        
        If Dir(szDir) <> "" Then
    
            nCount = nCount + 1
        Else
            Exit Do
        End If
    Loop
    nCount2 = 1
End Sub


Private Sub Timer2_Timer()
    Dim ctl As Control
    
    For Each ctl In Me
        If TypeOf ctl Is Label And ctl.Tag = "1" Then
         ctl.Visible = Not ctl.Visible
        End If
    Next
End Sub

