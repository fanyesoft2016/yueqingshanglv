VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form AppForm 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������̬������Ϣ"
   ClientHeight    =   600
   ClientLeft      =   3360
   ClientTop       =   2760
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "tpmsplay2.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":214E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":4152
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":4266
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":437A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":448E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":45A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":46B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":47CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tpmsplay2.frx":49F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "�½������ļ�"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "�������ļ�"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "����"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pause"
            Object.ToolTipText     =   "��ͣ"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "continue"
            Object.ToolTipText     =   "����"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous"
            Object.ToolTipText     =   "��һ��"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   "��һ��"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "ֹͣ"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "diswin"
            Object.ToolTipText     =   "ӳ�䴰��"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Object.ToolTipText     =   "�رղ��źʹ���"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "�˳�"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timerflic 
      Left            =   4800
      Top             =   2805
   End
   Begin VB.Timer TimerTXT 
      Enabled         =   0   'False
      Left            =   4785
      Top             =   2160
   End
   Begin VB.Timer TimerRTF 
      Left            =   5460
      Top             =   2775
   End
   Begin VB.CommandButton ReceiveKey 
      Height          =   285
      Left            =   -2000
      TabIndex        =   1
      Top             =   2985
      Width           =   345
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   4048
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "2016-7-5"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "15:20"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5475
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.Timer TimerPic 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5460
      Top             =   1605
   End
   Begin VB.Menu MENUFILE 
      Caption         =   "�ļ�(&F)"
      WindowList      =   -1  'True
      Begin VB.Menu NEWLED 
         Caption         =   "�½������ļ�LED..."
         Shortcut        =   ^E
      End
      Begin VB.Menu MODIFYLED 
         Caption         =   "�޸������ļ�LED..."
         Shortcut        =   ^M
      End
      Begin VB.Menu FILEOPEN 
         Caption         =   "��"
         Begin VB.Menu FILEPIC 
            Caption         =   "ͼƬ..."
         End
         Begin VB.Menu FILEAVI 
            Caption         =   "AVI�ļ�..."
         End
         Begin VB.Menu FILEFLIC 
            Caption         =   "FLIC����..."
         End
         Begin VB.Menu FILETXT 
            Caption         =   "��׼�ı�..."
         End
         Begin VB.Menu RTFFILE 
            Caption         =   "RTF�ļ�..."
         End
      End
      Begin VB.Menu SPACE4 
         Caption         =   "-"
      End
      Begin VB.Menu REALTIMEPLAY 
         Caption         =   "ʵʱ������Ʊ״����ʾ"
         Shortcut        =   ^D
      End
      Begin VB.Menu ASPACE 
         Caption         =   "-"
      End
      Begin VB.Menu FILELED 
         Caption         =   "�� &LED..."
         Shortcut        =   ^L
      End
      Begin VB.Menu SPACE1 
         Caption         =   "-"
      End
      Begin VB.Menu FILEEXIT 
         Caption         =   "�˳�    "
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu EDIT 
      Caption         =   "�༭��(&E)"
      Begin VB.Menu EDIT1 
         Caption         =   "�����ļ�(.LED)�༭��..."
      End
      Begin VB.Menu SPACE3 
         Caption         =   "-"
      End
      Begin VB.Menu EDIT2 
         Caption         =   "�ı��ļ�(.TXT)�༭��..."
      End
      Begin VB.Menu PBRUSH 
         Caption         =   "����..."
      End
   End
   Begin VB.Menu MENU3 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu TOOLBAR 
         Caption         =   "������"
         Checked         =   -1  'True
      End
      Begin VB.Menu STATUS 
         Caption         =   "״̬��"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MENU4 
      Caption         =   "ϵͳ����(&S)"
      Begin VB.Menu MENU41 
         Caption         =   "����ӳ��..."
      End
      Begin VB.Menu RTFOPTION 
         Caption         =   "RTF�ļ�����..."
         Visible         =   0   'False
      End
      Begin VB.Menu REALTIMEPLAYOPTIONS 
         Caption         =   "ʵʱ������Ʊ״����ʾ����..."
      End
      Begin VB.Menu SetTitle 
         Caption         =   "������ʾ������..."
      End
   End
   Begin VB.Menu MENU6 
      Caption         =   "��������(&P)"
      Begin VB.Menu PLAY 
         Caption         =   "����"
      End
      Begin VB.Menu PAUSE 
         Caption         =   "��ͣ"
      End
      Begin VB.Menu CONTINUE 
         Caption         =   "����"
      End
      Begin VB.Menu PREVIOUS 
         Caption         =   "��һ��"
      End
      Begin VB.Menu NEXT 
         Caption         =   "��һ��"
      End
      Begin VB.Menu STOP 
         Caption         =   "ֹͣ"
      End
      Begin VB.Menu SPACE2 
         Caption         =   "-"
      End
      Begin VB.Menu DISWINDOW 
         Caption         =   "ӳ�䴰��"
      End
      Begin VB.Menu CLOSEDISWIN 
         Caption         =   "�رղ��źʹ���"
      End
   End
   Begin VB.Menu MENU5 
      Caption         =   "����(&H)"
      Begin VB.Menu MENU51 
         Caption         =   "���ڲ������"
      End
   End
End
Attribute VB_Name = "AppForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FlagofPause As Boolean
Dim SumLengthofRTF As Long
Dim TXTPlayMode As Integer
Dim SumOfFrame As Integer
Dim FrameOfNumber As Integer
Dim flagofplaybarPressed As Boolean
Dim flagofStart As Boolean
Dim PauseTime As Integer
Dim sUserName As String
Dim sPassword As String
'½����ӵ�
Dim m_szServerName As String

Public Sub LoadRTFOption()
Dim a As Variant
Dim b As Variant
filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "rtf.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Close #filenum
RTFStartPosition = a
RTFStep = b
Exit Sub
dealerror:
Close #filenum
RTFStartPosition = 80
RTFStep = 10
End Sub

Public Sub SendNextMessage()
'AppForm.ReceiveKey.SetFocus
'SendKeys "prtsc"
AppForm.Caption = "���Ҽ�Ʊ���" + " -" + SavedLEDFile + " -" + RTrim(RecordArray(NumberofCurrentFile).filename)
PlayMulti
End Sub

Public Sub LoadRealtimePlayOption()
'������ʾ������

Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant
Dim filenum As Integer
filenum = FreeFile
On Error GoTo dealerror
Open CurrentPath + "database.cfg" For Input As #filenum
Input #filenum, a
Input #filenum, b
Input #filenum, c
Input #filenum, d
Close #filenum

iScrollRowofRealTime = a
iIntervaltimeofRealTime = b
sUserName = c
sPassword = d
Exit Sub
dealerror:
Close #filenum
iScrollRowofRealTime = 5
iIntervaltimeofRealTime = 5
sUserName = "centertest"
sPassword = ""
End Sub

Public Sub CloseRTF()
    TimerRTF.Enabled = False
    DisForm.RichTextBox1.width = 2
    DisForm.RichTextBox1.height = 2
    DisForm.RichTextBox1.Left = -30
End Sub

Public Sub OpenLEDFile(filename As String)
Dim LengthofRecord As Integer
Dim temprecord As RecordType
LengthofRecord = Len(temprecord)
FileLength& = FileLen(filename)
SumofRecords = FileLength / LengthofRecord
ReDim RecordArray(SumofRecords) As RecordType
Dim i As Integer
For i = 0 To SumofRecords
    RecordArray(i).filename = Space(100)
    RecordArray(i).TXTFontName = Space(30)
Next i
filenum% = FreeFile
'On Error GoTo dealerror
Open filename For Random As #filenum Len = LengthofRecord
For i% = 0 To SumofRecords - 1 Step 1
 Get #filenum, i + 1, RecordArray(i)
Next i
Close #filenum
SavedLEDFile = filename
Flag = True
'dealerror:
'action = MsgBox("������ļ���ʽ��", vbInformation, "��Ϣ")
End Sub




Public Sub DisplayWindow()
If flagofplaybarPressed = True Then
        Exit Sub
End If
DisForm.Left = gleft
DisForm.Top = gtop
DisForm.width = gwidth
DisForm.height = gheight
DisForm.ScaleWidth = gwidth * TwipspercentPixelX
DisForm.ScaleHeight = gheight * TwipspercentPixelY
DisForm.Show
End Sub
Public Sub PlayRTF()
'DisForm.MCIWnd1.Command = "close"
On Error GoTo dealerror
DisForm.RichTextBox1.Visible = False
DisForm.RichTextBox1.Left = 0
DisForm.RichTextBox1.Top = 0
DisForm.RichTextBox1.width = gwidth
DisForm.RichTextBox1.height = gheight
DisForm.RichTextBox1.LoadFile (RTrim(RecordArray(NumberofCurrentFile).filename))
DisForm.RichTextBox1.SelStart = 1000000
SumLengthofRTF = DisForm.RichTextBox1.SelStart
DisForm.RichTextBox1.SelStart = 0
DisForm.RichTextBox1.Visible = True
DisForm.RichTextBox1.SelStart = DisForm.RichTextBox1.SelStart + RTFStartPosition
DisForm.RichTextBox1.SelLength = 0
TimerRTF.Enabled = True
Exit Sub
dealerror:
action% = MsgBox("��RTF�ļ���ʽ����", vbInformation, "��Ϣ")
TimerRTF.Enabled = False
SendNextMessage
End Sub
Public Sub PlayPIC()
DisForm.Image1.Left = 0
DisForm.Image1.Top = 0
DisForm.Image1.height = gheight
DisForm.Image1.width = gwidth
DisForm.Image1.Picture = LoadPicture(RTrim(RecordArray(NumberofCurrentFile).filename))
End Sub
Private Sub PlayMulti()
If Flag = False Then
   'action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
   'CloseDISWindow
   Exit Sub
End If
filenum% = FreeFile
Select Case RecordArray(NumberofCurrentFile).Types
Case 3
   On Error GoTo dealerror3
   Open RTrim(RecordArray(NumberofCurrentFile).filename) For Input As #filenum
   Close #filenum
   
   If FirstFlags = True Then
      currentloops = 1
      FirstFlags = False
   End If
 
   DisForm.MCIWnd1.Speed = RecordArray(NumberofCurrentFile).Speed * 100
   DisForm.MCIWnd1.filename = RTrim(RecordArray(NumberofCurrentFile).filename)
   avilength = DisForm.MCIWnd1.End
   DisForm.MCIWnd1.Left = 0
   DisForm.MCIWnd1.Top = 0
   DisForm.MCIWnd1.width = gwidth
   DisForm.MCIWnd1.height = gheight
   DisForm.MCIWnd1.Command = "open"
   DisForm.MCIWnd1.Command = "Play"
  
   oldNumberofCurrentFile = NumberofCurrentFile
   
   If NumberofCurrentFile = SumofRecords - 1 Then
      NumberofCurrentFile = 0
   Else
      NumberofCurrentFile = NumberofCurrentFile + 1
   End If
   Exit Sub
dealerror3:
   PlayErrorDeal
   Close #filenum
   Exit Sub
Case 0
   On Error GoTo dealerror0
   Open RTrim(RecordArray(NumberofCurrentFile).filename) For Input As #filenum
   Close #filenum
   If FirstFlags = True Then
      currentloops = 1
      FirstFlags = False
    End If
   DisForm.Image1.width = 2
   DisForm.Image1.height = 2
   DisForm.Image1.Left = -20
   DisForm.MCIWnd1.width = 2
   DisForm.MCIWnd1.height = 2
   DisForm.MCIWnd1.Left = -20
   DisForm.RichTextBox1.width = 2
   DisForm.RichTextBox1.height = 2
   DisForm.RichTextBox1.Left = -20
   
   SumOfFrame = GetFrameNum(RTrim(RecordArray(NumberofCurrentFile).filename), DisForm.hwnd)
   FrameOfNumber = 0
   
   tempspeed% = (20 - RecordArray(NumberofCurrentFile).Speed) * 50
   If tempspeed = 0 Or tempspeed = 500 Then
      tempspeed = 1
   End If
   Timerflic.Interval = tempspeed
 
   oldNumberofCurrentFile = NumberofCurrentFile
   
   If NumberofCurrentFile = SumofRecords - 1 Then
      NumberofCurrentFile = 0
   Else
      NumberofCurrentFile = NumberofCurrentFile + 1
   End If
   
   Timerflic.Enabled = True
   Exit Sub
dealerror0:
   PlayErrorDeal
   Close #filenum
   Exit Sub
Case 1
    On Error GoTo dealerror1
   Open RTrim(RecordArray(NumberofCurrentFile).filename) For Input As #filenum
   Close #filenum
    If FirstFlags = True Then
      currentloops = 1
      FirstFlags = False
    End If
    TimerPic.Interval = Format(RecordArray(NumberofCurrentFile).Speed) * 1000
    TimerPic.Enabled = True
    
    Call PlayPIC
   
   oldNumberofCurrentFile = NumberofCurrentFile
    
    If NumberofCurrentFile = SumofRecords - 1 Then
       NumberofCurrentFile = 0
    Else
       NumberofCurrentFile = NumberofCurrentFile + 1
    End If
    Exit Sub
dealerror1:
   PlayErrorDeal
   Close #filenum
   Exit Sub
 Case 4
    On Error GoTo dealerror4
    Open RTrim(RecordArray(NumberofCurrentFile).filename) For Input As #filenum
    Close #filenum
   ' If FirstFlags = True Then
   '   currentloops = 1
    '  FirstFlags = False
    'End If
    
    TimerRTF.Interval = Format(RecordArray(NumberofCurrentFile).Speed) * 1000
    Call PlayRTF
    
    oldNumberofCurrentFile = NumberofCurrentFile
    
    If NumberofCurrentFile = SumofRecords - 1 Then
       NumberofCurrentFile = 0
    Else
       NumberofCurrentFile = NumberofCurrentFile + 1
    End If
    Exit Sub
dealerror4:
   PlayErrorDeal
   Close #filenum
   Exit Sub
 Case 2
    On Error GoTo dealerror2
    Open RTrim(RecordArray(NumberofCurrentFile).filename) For Input As #filenum
    Close #filenum
    If FirstFlags = True Then
      currentloops = 1
      FirstFlags = False
    End If
    DisForm.Image1.width = 2
    DisForm.Image1.height = 2
    DisForm.Image1.Left = -20
    DisForm.MCIWnd1.width = 2
    DisForm.MCIWnd1.height = 2
    DisForm.MCIWnd1.Left = -20
    DisForm.ForeColor = RecordArray(NumberofCurrentFile).TXTForeColor
    DisForm.FontSize = RecordArray(NumberofCurrentFile).TXTFontSize
    DisForm.FontBold = RecordArray(NumberofCurrentFile).TXTFontBold
    DisForm.FontItalic = RecordArray(NumberofCurrentFile).TXTFontItalic
    DisForm.FontUnderline = RecordArray(NumberofCurrentFile).TXTFontUnderline
    'DisForm.Font.Name = RecordArray(NumberofCurrentFile).TXTFontName
    'tempfont2 = RTrim(RecordArray(NumberofCurrentFile).filename)
    'tempfont1 = RTrim(RecordArray(NumberofCurrentFile).TXTFontName)
    'RecordArray(NumberofCurrentFile).TXTFontName = "sdfsdf"
   ' tempfontname = RTrim(RecordArray(NumberofCurrentFile).TXTFontName)
    DisForm.FontName = RTrim(RecordArray(NumberofCurrentFile).TXTFontName)
    TXTPlayMode = RecordArray(NumberofCurrentFile).TXTMode
    
    TimerTXT.Interval = Format(RecordArray(NumberofCurrentFile).Speed) * 1000
    PauseTime = Format(RecordArray(NumberofCurrentFile).Speed)
    PreTXTDo (DisForm.hdc)
    display_txt (RTrim(RecordArray(NumberofCurrentFile).filename))
    
    oldNumberofCurrentFile = NumberofCurrentFile
    
    If NumberofCurrentFile = SumofRecords - 1 Then
       NumberofCurrentFile = 0
    Else
       NumberofCurrentFile = NumberofCurrentFile + 1
    End If
    
    
    DisTXTLineORPingFirst (TXTPlayMode)
    
   
        
    'TimerTXT.Enabled = True
    'If FirstFlags = True Then
    '  currentloops = 1
    '  FirstFlags = False
    'End If
   
    Exit Sub
dealerror2:
   PlayErrorDeal
   Close #filenum
   Exit Sub
End Select
action = MsgBox("������ļ���ʽ��", 16, "��Ϣ")
If NumberofCurrentFile = SumofRecords - 1 Then
      NumberofCurrentFile = 0
Else
     NumberofCurrentFile = NumberofCurrentFile + 1
End If
SendNextMessage
End Sub
Public Sub NextMulti()
Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
        TimerRTF.Enabled = False
        CloseRTF
Case 1
        TimerPic.Enabled = False
Case 2
        TimerTXT.Enabled = False
        Close #TXTfilenum
Case 3
        DisForm.MCIWnd1.Command = "close"
Case 0
        Timerflic.Enabled = False
        action = Flcstop(AppForm.hwnd)
End Select
FirstFlags = True
SendNextMessage

End Sub
Public Sub PreviousMulti()
If NumberofCurrentFile > 1 Then
    NumberofCurrentFile = NumberofCurrentFile - 2
Else
    NumberofCurrentFile = 0
End If
Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
        TimerRTF.Enabled = False
        CloseRTF
Case 1
        TimerPic.Enabled = False
Case 2
        TimerTXT.Enabled = False
        Close #TXTfilenum
Case 3
        DisForm.MCIWnd1.Command = "close"
Case 0
        Timerflic.Enabled = False
        action = Flcstop(AppForm.hwnd)
End Select
FirstFlags = True
SendNextMessage
End Sub

Private Sub CARPLAY_Click()

End Sub

Private Sub CLOSEDISWIN_Click()
If Flag = False Then
   flagofdatabaseform = False
   Unload DisForm
   Unload DatabaseForm
   Exit Sub
End If

Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
    TimerRTF.Enabled = False
    CloseRTF
Case 1
    TimerPic.Enabled = False
    DisForm.Image1.Picture = LoadPicture("")
Case 2
    PauseTime = 0
    TimerTXT.Enabled = False
    Close #TXTfilenum
Case 3
    'temp$ = LCase$(Right$(RTrim(RecordArray(oldnumberofcurrentfil).filename), 3))
    'If temp = "dat" Then
    '   If flagofplaybarPressed = True And Flag = True Then
    '        DisForm.MCIWnd1.Command = "pause"
    '        action = MsgBox("����ֹͣ����VCD��", vbInformation, "��Ϣ")
    '        DisForm.MCIWnd1.Command = "resume"
    '        Exit Sub
    '   End If
    'End If
    DisForm.MCIWnd1.Command = "close"
Case 0
    Timerflic.Enabled = False
    action = Flcstop(AppForm.hwnd)
End Select

NumberofCurrentFile = 0
FirstFlags = True
Flag = False
flagofplaybarPressed = False

AppForm.Caption = "���Ҽ�Ʊ���"
CommonDialog1.filename = ""
If flagofdatabaseform = True Then
 flagofdatabaseform = False
 Unload DatabaseForm
End If
Unload DisForm
End Sub
Private Sub CONTINUE_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
   Exit Sub
End If
If FlagofPause = False Then
   Exit Sub
End If
Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
    FlagofPause = False
    TimerRTF.Enabled = True
    Exit Sub
Case 1
    FlagofPause = False
    TimerPic.Enabled = True
    Exit Sub
Case 2
    FlagofPause = False
    TimerTXT.Enabled = True
    Exit Sub
Case 3
    FlagofPause = False
    DisForm.MCIWnd1.Command = "resume"
    Exit Sub
Case 0
    FlagofPause = False
    FirstFlags = True
    SendNextMessage
    Exit Sub
End Select
End Sub

Private Sub DISWINDOW_Click()
Dim itest As Integer
itest = 2 + 1
If flagofdatabaseform = True Then
   Exit Sub
End If
DisplayWindow
End Sub

Private Sub EDIT1_Click()
NEWLED_Click
End Sub

Private Sub EDIT2_Click()
On Error GoTo openerror
returnCode = Shell("notepad.exe", vbNormalFocus)
Exit Sub
openerror:
  err
  action = MsgBox("��ȷ��notepad.exe������ڻ���ȷ��", 16, "����")
End Sub

Private Sub FILEAVI_Click()
Dim action As Integer
If Flag = True Then
   action = MsgBox("���ȹرղ�����ӳ�䴰��!", 16, "��Ϣ")
   Exit Sub
End If
CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "AVI�ļ�(*.avi)|*.avi|�����ļ�(*.*)|*.*"
          
CommonDialog1.ShowOpen
On Error GoTo openerror
If CommonDialog1.filename = "" Then
   Exit Sub
Else
DisplayWindow

DisForm.MCIWnd1.Command = "close"
DisForm.RichTextBox1.width = 2
DisForm.RichTextBox1.height = 2
DisForm.RichTextBox1.Left = -20
DisForm.Image1.width = 2
DisForm.Image1.height = 2
DisForm.Image1.Left = -20


DisForm.MCIWnd1.filename = CommonDialog1.filename
avilength = DisForm.MCIWnd1.End

DisForm.MCIWnd1.Left = 0
DisForm.MCIWnd1.Top = 0
DisForm.MCIWnd1.width = gwidth
DisForm.MCIWnd1.height = gheight
DisForm.MCIWnd1.Command = "Play"

CommonDialog1.filename = ""
End If
Exit Sub
openerror:
  err
  action = MsgBox("��Ч���ļ����ͣ�", 16, "����")
  CommonDialog1.filename = ""
End Sub

Private Sub FILEEXIT_Click()
'If flagofplaybarPressed = True Then
'    temp$ = LCase$(Right$(RTrim(RecordArray(oldnumberofcurrentfil).filename), 3))
'    If temp = "dat" Then
'            DisForm.MCIWnd1.Command = "pause"
'            action = MsgBox("����ֹͣ����VCD��", vbInformation, "��Ϣ")
'            DisForm.MCIWnd1.Command = "resume"
'            Exit Sub
'    End If
'End If
'If SavedLEDFile <> "" Then
'   filenum% = FreeFile
'   Open CurrentPath + "saved.cfg" For Output As #filenum
'   Print #filenum, SavedLEDFile
'   Close #filenum
'End If
If flagofdatabaseform = True Then
 Unload DatabaseForm
End If
If Flag = True Then
  CLOSEDISWIN_Click
End If
End
Unload AppForm
End Sub

Private Sub FILEFLIC_Click()
Dim action As Integer
If Flag = True Then
   action = MsgBox("���ȹرղ�����ӳ�䴰��!", 16, "��Ϣ")
   Exit Sub
End If

CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "FLIC�ļ�(*.flc,*.fli)|*.fli;*.flc|FLC�ļ�(*.flc)|*.flc|FLI�ļ�(*.fli)|*.fli|�����ļ�(*.*)|*.*"
          
CommonDialog1.ShowOpen
On Error GoTo openerror
If CommonDialog1.filename = "" Then
   Exit Sub
Else
DisplayWindow
   
   
DisForm.RichTextBox1.width = 2
DisForm.RichTextBox1.height = 2
DisForm.RichTextBox1.Left = -20
   DisForm.Image1.width = 2
   DisForm.Image1.height = 2
   DisForm.Image1.Left = -20
   DisForm.MCIWnd1.Command = "close"
   DisForm.MCIWnd1.width = 2
   DisForm.MCIWnd1.height = 2
   DisForm.MCIWnd1.Left = -20
   Sum = GetFrameNum(CommonDialog1.filename, DisForm.hwnd)
   action = ShowFrame(DisForm.hwnd, DisForm.hdc, CommonDialog1.filename, 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY, 1)
   action = Flcstop(AppForm.hwnd)
   CommonDialog1.filename = ""
End If
Exit Sub
openerror:
  err
  action = MsgBox("��Ч���ļ����ͣ�", 16, "����")
  CommonDialog1.filename = ""
End Sub

Private Sub FILELED_Click()
Dim action As Integer

CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "�����ļ�(*.led)|*.led"
CommonDialog1.ShowOpen
On Error GoTo openerror
If CommonDialog1.filename = "" Then
   Exit Sub
Else
If Flag = True Then
   STOP_Click
End If
If flagofdatabaseform = True Then
   Unload DatabaseForm
End If
OpenLEDFile (CommonDialog1.filename)
AppForm.Caption = "���Ҽ�Ʊ���" + " -" + CommonDialog1.filename
CommonDialog1.filename = ""
action = MsgBox("LED�ļ��Ѽ��أ��벥�ţ�", vbInformation, "��Ϣ")
End If
Exit Sub
openerror:
  'Err
  action = MsgBox("��Ч���ļ����ͣ�", 16, "����")
  CommonDialog1.filename = ""
End Sub

Private Sub FILEPIC_Click()
Dim action As Integer
If Flag = True Then
   action = MsgBox("���ȹرղ�����ӳ�䴰��!", 16, "��Ϣ")
   Exit Sub
End If
CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "ͼƬ�ļ�(*.bmp,*.jpg,*,gif)|*.bmp;*.jpg;*.gif|BMP�ļ�(*.bmp)|*.bmp|JPG�ļ�(*.jpg)|*.jpg|GIF�ļ�(*.gif)|*.gif|�����ļ�(*.*)|*.*"
          
CommonDialog1.ShowOpen
On Error GoTo openerror
If CommonDialog1.filename = "" Then
   Exit Sub
Else
DisplayWindow

DisForm.RichTextBox1.width = 2
DisForm.RichTextBox1.height = 2
DisForm.RichTextBox1.Left = -20
DisForm.MCIWnd1.Command = "close"
DisForm.MCIWnd1.width = 2
DisForm.MCIWnd1.height = 2
DisForm.MCIWnd1.Left = -20

DisForm.MCIWnd1.Command = "close"
DisForm.Image1.Left = 0
DisForm.Image1.height = gheight
DisForm.Image1.width = gwidth
DisForm.Image1.Picture = LoadPicture(CommonDialog1.filename)
'DisForm.Image1.Left = -20
CommonDialog1.filename = ""
End If
Exit Sub
openerror:
  err
  action = MsgBox("��Ч���ļ����ͣ�", 16, "����")
  CommonDialog1.filename = ""
End Sub

Private Sub FILESAVE_Click()

End Sub

Private Sub FILESAVEAS_Click()

End Sub

Private Sub FILETXT_Click()
If Flag = True Then
   action = MsgBox("���ȹرղ�����ӳ�䴰��!", 16, "��Ϣ")
   Exit Sub
End If
CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "��׼�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*"
          
CommonDialog1.ShowOpen
On Error GoTo openerror
If CommonDialog1.filename = "" Then
   Exit Sub
Else
DisplayWindow

DisForm.RichTextBox1.width = 2
DisForm.RichTextBox1.height = 2
DisForm.RichTextBox1.Left = -20
DisForm.MCIWnd1.Command = "close"
DisForm.MCIWnd1.width = 2
DisForm.MCIWnd1.height = 2
DisForm.MCIWnd1.Left = -20
DisForm.Image1.width = 2
DisForm.Image1.height = 2
DisForm.Image1.Left = -20

TXTPlayMode = 1
    
PreTXTDo (DisForm.hdc)
display_txt (CommonDialog1.filename)
DisTXTLineORPingFirst (TXTPlayMode)
Close #TXTfilenum
CommonDialog1.filename = ""
End If
Exit Sub
openerror:
  'Err
  action = MsgBox("��Ч���ļ����ͣ�", 16, "����")
  CommonDialog1.filename = ""
End Sub

Private Sub FONT_Click()
CommonDialog1.ShowFont
End Sub

Private Sub Form_Activate()
If flagofStart = True Then
  If FlagofAuto = True Then
   LoadLEDFileandPlay
  Else
   LoadLEDFile
  End If
  flagofStart = False
End If
End Sub

Private Sub Form_Load()
'Myscreen = New Button
'MMControl1.hWndDisplay = Picture1.hWnd
'MMControl1.Notify = False
 '   MMControl1.Wait = True
  '  MMControl1.Shareable = False
   ' MMControl1.DeviceType = "AVIVideo"
  '  MMControl1.filename = "d:\picture\flc\welcome1.avi"

    '�� MCI WaveAudio �豸��  single
'    MMControl1.Command = "Open"
'action = MsgBox(App.Path, 16, "information")

'Ӧ�ó���ʼ����¼
'Set oGate = CreateObject("prjGate.Gate")
'Set oUser = oGate.Enter("zw", "zw", "", "")
'����˵����
'�û���
'����
'IP��ַ:��:192.168.1.51
'������:��:F7823478
Dim szConnection As String


oTo.ConnectionString = GetConnectionStr("")

CurrentPath = App.Path + "\"


CommonDialog1.FontName = "����"
CommonDialog1.FontSize = 11
CommonDialog1.Color = RGB(255, 0, 0)
CommonDialog1.FontBold = False
CommonDialog1.FontItalic = False
CommonDialog1.FontUnderline = False

TimerPic.Enabled = False
TimerRTF.Enabled = False
TimerTXT.Enabled = False
flagofModify = False
flagofStart = True

flagofplaybarPressed = False
Flag = False
FirstFlags = True
FlagofPause = False
FlagofAuto = False


LoadRTFOption

'������ʾ������
LoadRealtimePlayOption

TwipspercentPixelX = Screen.TwipsPerPixelX
TwipspercentPixelY = Screen.TwipsPerPixelY
'
NumberofCurrentFile = 0

LoadCoordOption
'
AppForm.Left = 2000

AppForm.Top = gtop + gheight

On Error GoTo here
m_szServerName = GetServerName
oUser.Login sUserName, sPassword, m_szServerName
oSystem.Init oUser

'�õ�������ʱ��
Date = oSystem.NowDate
Time = oSystem.NowTime
'oSellTicketClient.Init oUser

'Ӧ�ó���ʼ����¼
'Set oGate = CreateObject("prjGate.Gate")
''Set oUser = oGate.enter("yxd", "dong", "", "")
'Set oUser = oGate.Enter(sUserName, sPassword, "", "")
'Me.Show

REALTIMEPLAY_Click


Exit Sub
here:
'MsgBox Err.Description, vbInformation, Err.Number
Dim action As Integer
If err.Number = 2253 Then
    action = MsgBox("���ܵ�¼���������޴��û���", vbInformation, "����")
  Else
    action = MsgBox("��¼�������������ܽ���ʵʱ������Ʊ״����ʾ��", vbInformation, "����")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'action = MsgBox("exit the program", 16, "informatio")

If SavedLEDFile <> "" Then
   filenum% = FreeFile
   Open CurrentPath + "saved.cfg" For Output As #filenum
   Print #filenum, SavedLEDFile
   Close #filenum
End If
If flagofdatabaseform = True Then
 Unload DatabaseForm
End If
If Flag = True Then
  CLOSEDISWIN_Click
  Exit Sub
End If
End
End Sub

Private Sub MENU41_Click()
If flagofplaybarPressed = True Then
 action = MsgBox("����ֹͣ���ţ��ٽ��е�����", vbInformation, "��Ϣ")
 Exit Sub
End If
CoordForm.Show
End Sub



Private Sub Modify_Click()
Dim action As Integer
CommonDialog1.DialogTitle = "��Ҫ�޸ĵ������ļ�"
CommonDialog1.Filter = "�����ļ�(*.led)|*.led"
          
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then
   Exit Sub
Else
  sTempFilename = CommonDialog1.filename
  flagofModify = True
  LedForm.Caption = "�޸������ļ�" + " -" + CommonDialog1.filename
  CommonDialog1.filename = ""
  LedForm.Show
End If
End Sub

Private Sub MENU51_Click()
ShowAbout
End Sub

Private Sub mnu_SpringTime_Click()
    
End Sub

Private Sub MODIFYLED_Click()
Dim action As Integer
CommonDialog1.DialogTitle = "��Ҫ�޸ĵ������ļ�"
CommonDialog1.Filter = "�����ļ�(*.led)|*.led"
          
CommonDialog1.ShowOpen

If CommonDialog1.filename = "" Then
   Exit Sub
Else
  sTempFilename = CommonDialog1.filename
  flagofModify = True
  LedForm.Caption = "�޸������ļ�" + " -" + CommonDialog1.filename
  CommonDialog1.filename = ""
  LedForm.Show 1
End If
End Sub

Private Sub NEWLED_Click()
LedForm.Show 1
End Sub

Private Sub NEXT_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
      Exit Sub
End If
If FlagofPause = True Then
   FlagofPause = False
End If

NextMulti

End Sub

Private Sub PBRUSH_Click()
On Error GoTo openerror
returnCode = Shell("pbrush.exe", vbNormalFocus)
Exit Sub
openerror:
  err
  action = MsgBox("��ȷ��pbrush.exe������ڻ���ȷ��", 16, "����")
End Sub

Private Sub PLAY_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
   Exit Sub
End If
If flagofplaybarPressed = False Then
    DisplayWindow
    flagofplaybarPressed = True
    SendNextMessage

Else
    Exit Sub
End If
End Sub

Private Sub PREVIOUS_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
   Exit Sub
End If
If FlagofPause = True Then
   FlagofPause = False
End If
PreviousMulti

End Sub

Private Sub PAUSE_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
      Exit Sub
End If
If FlagofPause = True Then
   Exit Sub
End If
Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
    TimerRTF.Enabled = False
Case 1
    TimerPic.Enabled = False
Case 2
    TimerTXT.Enabled = False
Case 3
    DisForm.MCIWnd1.Command = "pause"
Case 0
    Timerflic.Enabled = False
    action = Flcstop(AppForm.hwnd)
End Select
FlagofPause = True
End Sub



Private Sub REALTIMEPLAY_Click()
If flagofplaybarPressed = True Then
      CLOSEDISWIN_Click
End If
Flag = False
flagofdatabaseform = True
AppForm.Caption = "���Ҽ�Ʊ���"
DatabaseForm.Left = gleft
DatabaseForm.Top = gtop
DatabaseForm.width = gwidth
DatabaseForm.height = gheight
DatabaseForm.Show
End Sub

Private Sub REALTIMEPLAYOPTIONS_Click()
RealTimeForm.Show 1
End Sub

Private Sub ReceiveKey_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("prtsc") Then
   PlayMulti
End If
End Sub

Private Sub RTFFILE_Click()
If Flag = True Then
   action = MsgBox("���ȹرղ�����ӳ�䴰��!", 16, "��Ϣ")
   Exit Sub
End If
CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "RTF�ļ�(*.RTF)|*.RTF"
          
CommonDialog1.ShowOpen
If CommonDialog1.filename = "" Then
   Exit Sub
Else
DisplayWindow
DisForm.MCIWnd1.Command = "close"
DisForm.MCIWnd1.width = 2
DisForm.MCIWnd1.height = 2
DisForm.MCIWnd1.Left = -20
DisForm.Image1.width = 2
DisForm.Image1.height = 2
DisForm.Image1.Left = -20

DisForm.RichTextBox1.Top = 0
DisForm.RichTextBox1.Left = 0
DisForm.RichTextBox1.height = gheight
DisForm.RichTextBox1.width = gwidth
DisForm.RichTextBox1.LoadFile (CommonDialog1.filename)
CommonDialog1.filename = ""
End If
End Sub

Private Sub RTFOPTION_Click()
RTFForm.Show 1
End Sub

Private Sub SetTitle_Click()
    frmSetTitle.Show vbModal
End Sub

Private Sub STATUS_Click()
If STATUS.Checked = True Then
   STATUS.Checked = False
   StatusBar1.Visible = False
Else
   STATUS.Checked = True
   StatusBar1.Visible = True
End If
End Sub

Private Sub STOP_Click()
If Flag = False Then
   action = MsgBox("���ȴ������ļ�(.LED)!", 16, "��Ϣ")
      Exit Sub
End If
NumberofCurrentFile = 0
FirstFlags = True
flagofplaybarPressed = False
Select Case RecordArray(oldNumberofCurrentFile).Types
Case 4
    TimerRTF.Enabled = False
    CloseRTF
Case 1
    TimerPic.Enabled = False
    DisForm.Image1.Picture = LoadPicture("")
Case 2
    'Break
    PauseTime = 0
    TimerTXT.Enabled = False
    Close #TXTfilenum
Case 3
    DisForm.MCIWnd1.Command = "close"
Case 0
    Timerflic.Enabled = False
    action = Flcstop(AppForm.hwnd)
End Select

End Sub

Private Sub Timerflic_Timer()
If FrameOfNumber < SumOfFrame - 1 Then
   result = ShowFrame(DisForm.hwnd, DisForm.hdc, RTrim(RecordArray(oldNumberofCurrentFile).filename), 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY, FrameOfNumber)
   FrameOfNumber = FrameOfNumber + 1
Else
   result = ShowFrame(DisForm.hwnd, DisForm.hdc, RTrim(RecordArray(oldNumberofCurrentFile).filename), 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY, FrameOfNumber)
   action = Flcstop(AppForm.hwnd)
   Timerflic.Enabled = False
   If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         'action = MsgBox(Str$(currentloops), 16, "information")
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
   Else
        FirstFlags = True
   End If
   SendNextMessage
End If
End Sub

Private Sub timerPic_Timer()
TimerPic.Enabled = False
If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         'action = MsgBox(Str$(currentloops), 16, "information")
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
Else
        FirstFlags = True
End If
SendNextMessage
End Sub

Private Sub timerRTF_Timer()
If DisForm.RichTextBox1.SelStart = SumLengthofRTF Then
    TimerRTF.Enabled = False
    If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         'action = MsgBox(Str$(currentloops), 16, "information")
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
    Else
         FirstFlags = True
    End If
    CloseRTF
    SendNextMessage
Else
    DisForm.RichTextBox1.SelStart = DisForm.RichTextBox1.SelStart + RTFStep
End If
End Sub

Private Sub TimerTXT_Timer()
DisTXTLineORPing (TXTPlayMode)
End Sub



Private Sub TOOLBAR_Click()
If TOOLBAR.Checked = True Then
   TOOLBAR.Checked = False
   Toolbar1.Visible = False
Else
   TOOLBAR.Checked = True
   Toolbar1.Visible = True
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "close"
     CLOSEDISWIN_Click
Case "diswin"
     DISWINDOW_Click
Case "continue"
      CONTINUE_Click
Case "play"
      PLAY_Click
Case "next"
      NEXT_Click
Case "stop"
      STOP_Click
Case "pause"
      PAUSE_Click
Case "previous"
      PREVIOUS_Click
Case "open"
      FILELED_Click
Case "new"
      NEWLED_Click
Case "exit"
      FILEEXIT_Click
End Select
End Sub

Public Sub DisTXTLineORPing(ByVal mode As Integer)
Dim Buffer As String * 1
Dim recttemp As RECT
Dim rect2 As RECT
Dim rTemp As RECT
action = SetRect(rTemp, 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY)
action = SetRect(rect2, 0, gheight / TwipspercentPixelY - TXTfh, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY)
Select Case mode
 Case 0
 clounm% = 0
 Row% = gheight / TwipspercentPixelX - TXTfh
 action = ScrollDC(DisForm.hdc, 0, -TXTfh, rTemp, rTemp, 0, recttemp)
 action = FillRect(DisForm.hdc, rect2, hbr)
 Do While Not EOF(TXTfilenum)
   Get #TXTfilenum, , Buffer
   If Asc(Buffer) = &HA Then
       Exit Sub
   End If
      
   action = TextOut(DisForm.hdc, clounm, Row, Buffer, 2)
  
   clounm = clounm + TXTfw
   If clounm + TXTfw > gwidth / TwipspercentPixelX Then
        Exit Sub
   End If
 Loop
  
 
 TimerTXT.Enabled = False
 action = DeleteObject(hbr)
 Close #TXTfilenum
 
 If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         'action = MsgBox(Str$(currentloops), 16, "information")
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
 Else
         FirstFlags = True
 End If
 
' Start# = Timer   ' ���ÿ�ʼ��ͣ��ʱ�̡�
'    Do While Timer < Start + PauseTime
'        DoEvents    ' �������ø���������
'    Loop
 If Flag = False Or flagofplaybarPressed = False Then
     Exit Sub
 End If
 SendNextMessage
 Exit Sub
Case 1
 clounm% = 0
 Row% = 0
 action = FillRect(DisForm.hdc, rTemp, hbr)
 Do While Not EOF(TXTfilenum)
  Get #TXTfilenum, , Buffer
 
  If Asc(Buffer) = &HA Then
     'temp = Buffer
     clounm = gwidth / TwipspercentPixelY
     clounm = clounm + TXTfw
     If clounm + TXTfw > gwidth / TwipspercentPixelX Then
          Row = Row + TXTfh
          If Row + TXTfh <= gheight / TwipspercentPixelY Then
            clounm = 0
          Else
             Exit Sub
          End If
      End If
      
  End If
  
  action = TextOut(DisForm.hdc, clounm, Row, Buffer, 2)
  
  clounm = clounm + TXTfw
  If clounm + TXTfw > gwidth / TwipspercentPixelX Then
     Row = Row + TXTfh
     If Row + TXTfh <= gheight / TwipspercentPixelY Then
       clounm = 0
     Else
       If Flag = False Then
        action = DeleteObject(hbr)
        Close #TXTfilenum
       End If
       Exit Sub
     End If
  End If
 Loop
  
 
    
 TimerTXT.Enabled = False
 action = DeleteObject(hbr)
 Close #TXTfilenum
 
 
 If currentloops < RecordArray(oldNumberofCurrentFile).Loops Then
         'action = MsgBox(Str$(currentloops), 16, "information")
         NumberofCurrentFile = oldNumberofCurrentFile
         currentloops = currentloops + 1
 Else
         FirstFlags = True
 End If
 
'  Start# = Timer   ' ���ÿ�ʼ��ͣ��ʱ�̡�
'    Do While Timer < Start + PauseTime
'        DoEvents    ' �������ø���������
'    Loop
 If Flag = False Or flagofplaybarPressed = False Then
     Exit Sub
 End If
 SendNextMessage
 Exit Sub
End Select
End Sub

Public Sub display_txt(ByVal filename As String)

TXTfilenum = FreeFile
'TXTfilelength = FileLen(filename)
If TXTfw > gwidth / TwipspercentPixelX Or TXTfh > gheight / TwipspercentPixelY Then
         action = MsgBox("������������޷���ʾ��", vbInformation, "��Ϣ")
         Exit Sub
End If
On Error GoTo dealerror
Open filename For Binary As #TXTfilenum
If Flag = True Then
   TimerTXT.Enabled = True
End If
Exit Sub

dealerror:
action = MsgBox("���ļ���", vbInformation, "��Ϣ")
Exit Sub
End Sub
Public Sub PreTXTDo(ByVal hdc As Long)
  Dim txtsize As Size
  Dim lptm As TEXTMETRIC
  Dim rTemp As RECT
  action = SetBkColor(hdc, RGB(0, 0, 0))
  hbr = CreateSolidBrush(RGB(0, 0, 0))
  action = SelectObject(hdc, hbr)
  action = SetRect(rTemp, 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY)
      
      If GetTextColor(hdc) = RGB(0, 0, 0) Then
          action = SetTextColor(hdc, RGB(255, 255, 255))
      End If

      If GetTextMetrics(hdc, lptm) = 0 Then
        action = GetTextExtentPoint32(hdc, "AB", 2, txtsize)
        TXTfh = txtsize.cx
        TXTfw = txtsize.cy
      Else
       action = GetTextMetrics(hdc, lptm)
       TXTfw = lptm.tmMaxCharWidth
       TXTfh = lptm.tmHeight
      End If
      
   action = FillRect(hdc, rTemp, hbr)
End Sub
Public Sub LoadLEDFileandPlay()
Dim a As Variant
filenum% = FreeFile
On Error GoTo dealerror
Open CurrentPath + "saved.cfg" For Input As #filenum
Input #filenum, a
Close #filenum
SavedLEDFile = a
Open SavedLEDFile For Input As #filenum
Close #filenum
If SavedLEDFile <> "" Then
   OpenLEDFile SavedLEDFile
   AppForm.Caption = "���Ҽ�Ʊ���" + " -" + SavedLEDFile + " -" + RTrim(RecordArray(NumberofCurrentFile).filename)
   PLAY_Click
   Exit Sub
Else
   Exit Sub
End If
dealerror:
Close #filenum
Exit Sub
End Sub
Public Sub LoadCoordOption()
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
gleft = a * TwipspercentPixelX
gtop = b * TwipspercentPixelY
gwidth = c * TwipspercentPixelX
gheight = d * TwipspercentPixelY
FlagofAuto = e
Exit Sub
dealerror:
gleft = 0
gtop = 0
gwidth = 400 * TwipspercentPixelX
gheight = 300 * TwipspercentPixelY
FlagofAuto = 0
Close #filenum
End Sub

Public Sub LoadLEDFile()
'Dim a As Variant
'filenum% = FreeFile
'On Error GoTo dealerror
'Open CurrentPath + "saved.cfg" For Input As #filenum
'Input #filenum, a
'Close #filenum
'SavedLEDFile = a
'Open SavedLEDFile For Input As #filenum
'Close #filenum
'If SavedLEDFile <> "" Then
'   OpenLEDFile SavedLEDFile
'   AppForm.Caption = "���Ҽ�Ʊ���" + " -" + SavedLEDFile
'   Exit Sub
'Else
'   Exit Sub
'End If
'dealerror:
'Close #filenum
'Exit Sub
End Sub

Public Sub PlayErrorDeal()
  action% = MsgBox("��" + RTrim(RecordArray(NumberofCurrentFile).filename) + "�ļ�������ļ�������,�Ƿ�رղ��ţ������޸�?", vbYesNo, "��Ϣ")
  Select Case action
  Case 7
    If NumberofCurrentFile = SumofRecords - 1 Then
       NumberofCurrentFile = 0
    Else
       NumberofCurrentFile = NumberofCurrentFile + 1
    End If
    SendNextMessage
  Case 6
    CLOSEDISWIN_Click
   End Select
End Sub

Public Sub DisTXTLineORPingFirst(ByVal mode As Integer)
Dim Buffer As String * 1
Dim recttemp As RECT
Dim rect2 As RECT
Dim rTemp As RECT
action = SetRect(rTemp, 0, 0, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY)
action = SetRect(rect2, 0, gheight / TwipspercentPixelY - TXTfh, gwidth / TwipspercentPixelX, gheight / TwipspercentPixelY)
Select Case mode
Case 0
 clounm% = 0
 Row% = gheight / TwipspercentPixelX - TXTfh
 action = ScrollDC(DisForm.hdc, 0, -TXTfh, rTemp, rTemp, 0, recttemp)
 action = FillRect(DisForm.hdc, rect2, hbr)
 Do While Not EOF(TXTfilenum)
   Get #TXTfilenum, , Buffer
   If Asc(Buffer) = &HA Then
       Exit Sub
   End If
      
   action = TextOut(DisForm.hdc, clounm, Row, Buffer, 2)
  
   clounm = clounm + TXTfw
   If clounm + TXTfw > gwidth / TwipspercentPixelX Then
        Exit Sub
   End If
 Loop
  
 
 Exit Sub
Case 1
 clounm% = 0
 Row% = 0
 action = FillRect(DisForm.hdc, rTemp, hbr)
 Do While Not EOF(TXTfilenum)
  Get #TXTfilenum, , Buffer
 
  If Asc(Buffer) = &HA Then
     'temp = Buffer
     clounm = gwidth / TwipspercentPixelY
     clounm = clounm + TXTfw
     If clounm + TXTfw > gwidth / TwipspercentPixelX Then
          Row = Row + TXTfh
          If Row + TXTfh <= gheight / TwipspercentPixelY Then
            clounm = 0
          Else
             Exit Sub
          End If
      End If
      
  End If
  
  action = TextOut(DisForm.hdc, clounm, Row, Buffer, 2)
  
  clounm = clounm + TXTfw
  If clounm + TXTfw > gwidth / TwipspercentPixelX Then
     Row = Row + TXTfh
     If Row + TXTfh <= gheight / TwipspercentPixelY Then
       clounm = 0
     Else
       If Flag = False Then
        action = DeleteObject(hbr)
        Close #TXTfilenum
       End If
       Exit Sub
     End If
  End If
 Loop
 
 Exit Sub
End Select
End Sub

Public Sub RealTimePlayDatabasePackage()

End Sub
Private Function GetServerName() As String
    Dim oReg As New CFreeReg
    Dim szServer As String, szDatabaseType As String
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=foricq;Data Source=jhxu
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"  'HKEY_LOCAL_MACHINE
    '1�Ƚ�Ĭ��ֵ����
    szDatabaseType = oReg.GetSetting("DataBaseSet", "DBType")
    If szDatabaseType <> "" Then
        szServer = oReg.GetSetting("DataBaseSet", "DBServer")
    End If
    GetServerName = szServer
End Function


Private Sub ShowAbout()
'    Dim oShell As New CommShell
'    On Error GoTo ErrorHandle
'    oShell.ShowAbout "���Ҽ�Ʊ���", "Multimedia Electron Display Play Soft", "���Ҽ�Ʊ���", Me.Icon, App.Major, App.Minor, App.Revision
'    Set oShell = Nothing
'    Exit Sub
'ErrorHandle:
'    Set oShell = Nothing
End Sub
