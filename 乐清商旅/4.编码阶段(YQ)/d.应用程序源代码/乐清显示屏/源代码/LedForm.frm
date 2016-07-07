VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form LedForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "建立描述文件"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   6
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LedForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6750
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5070
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "字体..."
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
      Left            =   5445
      TabIndex        =   24
      Top             =   3195
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "颜色..."
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
      Left            =   3930
      TabIndex        =   23
      Top             =   3210
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "整屏翻页"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   22
      Top             =   3255
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "文本文件播放模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   3480
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "逐行上卷"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   225
      LargeChange     =   10
      Left            =   4350
      Max             =   60
      Min             =   1
      TabIndex        =   19
      Top             =   2760
      Value           =   1
      Width           =   2250
   End
   Begin VB.TextBox PTText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5805
      TabIndex        =   18
      Top             =   2355
      Width           =   780
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   225
      LargeChange     =   5
      Left            =   2370
      Max             =   20
      Min             =   1
      TabIndex        =   16
      Top             =   2745
      Value           =   10
      Width           =   1875
   End
   Begin VB.TextBox FAText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   300
      Left            =   3660
      TabIndex        =   15
      Top             =   2385
      Width           =   570
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "LedForm.frx":0442
      Left            =   2565
      List            =   "LedForm.frx":0444
      TabIndex        =   11
      Text            =   "*.flc"
      Top             =   1785
      Width           =   1440
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   100
      Left            =   120
      Max             =   10000
      Min             =   1
      TabIndex        =   10
      Top             =   2745
      Value           =   1
      Width           =   1860
   End
   Begin VB.TextBox Looptext 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1275
      TabIndex        =   9
      Top             =   2385
      Width           =   690
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
      Height          =   405
      Index           =   3
      Left            =   5865
      TabIndex        =   7
      Top             =   1665
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
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
      Index           =   2
      Left            =   5850
      TabIndex        =   6
      Top             =   1140
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
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
      Left            =   5835
      TabIndex        =   5
      Top             =   600
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
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
      Index           =   0
      Left            =   5835
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
   Begin VB.FileListBox File1 
      Archive         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2535
      TabIndex        =   3
      Top             =   450
      Width           =   1485
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4095
      TabIndex        =   2
      Top             =   900
      Width           =   1515
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4125
      TabIndex        =   1
      Top             =   465
      Width           =   1500
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "LedForm.frx":0446
      Left            =   45
      List            =   "LedForm.frx":0448
      TabIndex        =   0
      Top             =   75
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "图片,文本停留时间(秒)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4305
      TabIndex        =   17
      Top             =   2220
      Width           =   1440
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   135
      Width           =   2490
   End
   Begin VB.Label Label3 
      Caption         =   "文件："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2550
      TabIndex        =   13
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FLIC,AVI播放速度："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2355
      TabIndex        =   12
      Top             =   2205
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "循环次数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   8
      Top             =   2385
      Width           =   1275
   End
End
Attribute VB_Name = "LedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delindex As Integer
Dim SumofRecord As Integer
Dim Records(1000) As RecordType
Dim Sum As Integer
Dim TempArray() As RecordType

Public Sub SaveRecord()
Dim filenum As Integer
Dim LengthofRecord As Integer
Dim temprecord As RecordType
'Dim temp As SECURITY_ATTRIBUTES
If List1.ListCount > 1000 Then
   action = MsgBox("建立项目不能超过1000个！", 16, "消息")
   Exit Sub
End If
LengthofRecord = Len(temprecord)
CommonDialog1.DialogTitle = "另存为"
CommonDialog1.Filter = "描述文件(*.led)|*.led"
CommonDialog1.ShowSave
If CommonDialog1.filename = "" Then
   Exit Sub
End If
'action = MsgBox(tempfilename, vbInformation, "information")
filenum% = FreeFile
'Handle& = CreateFile(CommonDialog1.filename, GENERIC_READ, FILE_SHARE_READ, temp, CREATE_NEW, FILE_FLAG_RANDOM_ACCESS, 0)
On Error GoTo dealerror
Open CommonDialog1.filename For Input As #filenum
 'If Handle <> -1 Then
    action = MsgBox("此文件已存在，要覆盖吗？", vbYesNo, "警告")
    If action = 6 Then
        Close #filenum
        Kill CommonDialog1.filename
        Open CommonDialog1.filename For Random As #filenum Len = LengthofRecord
        For i% = 0 To List1.ListCount - 1 Step 1
          Put #filenum, i + 1, Records(i)
        Next i
        Close #filenum
        CommonDialog1.filename = ""
        Exit Sub
    Else
        Exit Sub
    End If
dealerror:
Open CommonDialog1.filename For Random As #filenum Len = LengthofRecord
For i% = 0 To List1.ListCount - 1 Step 1
  Put #filenum, i + 1, Records(i)
Next i
Close #filenum
CommonDialog1.filename = ""
End Sub
Private Sub Combo1_Click()
File1.Pattern = Combo1.Text

End Sub

Private Sub Command1_Click(Index As Integer)
Dim tempstring As String
Dim i As Integer
Dim temprecord As RecordType
'Dim recordstring As String
Dim temp As String
Select Case Index
Case 0
    If File1.filename = "" Then
       action = MsgBox("未选项目，不能添加！", vbInformation, "消息")
       Exit Sub
    End If
   SumofRecord = List1.ListCount
   If SumofRecord > 10000 Then
       action = MsgBox("最多不能超过1000项！", 16, "消息")
       Exit Sub
   End If
   temp = Right(File1.filename, 3)
   temp = LCase$(temp)
    Select Case temp
    Case "flc"
        tempstring = File1.filename + " -L" + Looptext.Text + "-FA" + FAText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 0
        Records(SumofRecord).Speed = Val(FAText.Text)
    Case "fli"
        tempstring = File1.filename + " -L" + Looptext.Text + "-FA" + FAText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 0
        Records(SumofRecord).Speed = Val(FAText.Text)
    Case "avi"
        tempstring = File1.filename + " -L" + Looptext.Text + "-FA" + FAText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 3
        Records(SumofRecord).Speed = Val(FAText.Text)
   'Case "dat"
   '     tempstring = File1.filename + " -L" + Looptext.Text + "-FA" + FAText.Text
   '     Records(SumofRecord).NumberofRecord = SumofRecord
   '     If Right(File1.Path, 1) = "\" Then
   '             Records(SumofRecord).filename = File1.Path + File1.filename
   '     Else
   '             Records(SumofRecord).filename = File1.Path + "\" + File1.filename
   '     End If
   '     Records(SumofRecord).Loops = Val(Looptext.Text)
   '     Records(SumofRecord).Types = 3
   '     Records(SumofRecord).Speed = Val(FAText.Text)
    Case "jpg"
        tempstring = File1.filename + " -L" + Looptext.Text + "-PT" + PTText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 1
        Records(SumofRecord).Speed = Val(PTText.Text)
    Case "bmp"
        tempstring = File1.filename + " -L" + Looptext.Text + "-PT" + PTText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 1
        Records(SumofRecord).Speed = Val(PTText.Text)
    Case "gif"
        tempstring = File1.filename + " -L" + Looptext.Text + "-PT" + PTText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 1
        Records(SumofRecord).Speed = Val(PTText.Text)
    Case "rtf"
        tempstring = File1.filename + " -L" + Looptext.Text + "-PT" + PTText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 4
        Records(SumofRecord).Speed = Val(PTText.Text)
   
   Case "txt"
        tempstring = File1.filename + " -L" + Looptext.Text + "-PT" + PTText.Text
        Records(SumofRecord).NumberofRecord = SumofRecord
        If Right(File1.Path, 1) = "\" Then
                Records(SumofRecord).filename = File1.Path + File1.filename
        Else
                Records(SumofRecord).filename = File1.Path + "\" + File1.filename
        End If
        Records(SumofRecord).Loops = Val(Looptext.Text)
        Records(SumofRecord).Types = 2
        Records(SumofRecord).Speed = Val(PTText.Text)
        If Option1(0) = True Then
          Records(SumofRecord).TXTMode = 0
        Else
          Records(SumofRecord).TXTMode = 1
        End If
        Records(SumofRecord).TXTFontBold = AppForm.CommonDialog1.FontBold
        Records(SumofRecord).TXTFontItalic = AppForm.CommonDialog1.FontItalic
        Records(SumofRecord).TXTFontUnderline = AppForm.CommonDialog1.FontUnderline
        Records(SumofRecord).TXTFontName = AppForm.CommonDialog1.FontName
        Records(SumofRecord).TXTFontSize = AppForm.CommonDialog1.FontSize
        Records(SumofRecord).TXTForeColor = AppForm.CommonDialog1.Color
   'Case "led"
   '    Dim LengthofRecord As Integer
   '    LengthofRecord = Len(temprecord)
   '    filelength& = FileLen(File1.filename)
   '    tempSumofRecords% = filelength / LengthofRecord
   '    filenum% = FreeFile
   '    Open File1.filename For Random As #filenum Len = LengthofRecord
   '    For i% = 0 To tempSumofRecords - 1 Step 1
   '    Get #filenum, i + 1, Records(i)
   '    List1.AddItem Records(i).filename
   '    Next i
   '    Close #filenum
       
    End Select
    'recordstring = File1.Path + "\" + tempstring
    List1.AddItem tempstring
   ' action = MsgBox(Str$(SumofRecord), 16, "sumforecord")

Case 1
    If List1.ListCount < 1 Then
       action = MsgBox("项目已删除完毕！", vbInformation, "消息")
       Exit Sub
    End If
    If List1.ListIndex = -1 Then
       action = MsgBox("未选中要删除项目！", vbInformation, "消息")
       Exit Sub
    End If
    
    For i = List1.ListIndex To List1.ListCount - 1 Step 1
       Records(List1.ListIndex) = Records(List1.ListIndex + 1)
    Next i

    List1.RemoveItem (List1.ListIndex)
   ' action = MsgBox(Str$(SumofRecord), 16, "sumforecord")
    List1.SetFocus
Case 2
    If List1.ListCount = 0 Then
        action = MsgBox("未编辑项目！", vbInformation, "消息")
        Exit Sub
    End If
    SaveRecord
    'Unload LedForm
Case 3
    Unload LedForm
End Select
End Sub

Private Sub Command2_Click()
AppForm.CommonDialog1.ShowColor
End Sub

Private Sub Command3_Click()
AppForm.CommonDialog1.ShowFont
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Label4.Caption = File1.Path + "\" + File1.filename
End Sub

Private Sub File1_DblClick()
Command1_Click (0)
End Sub

Private Sub Form_Load()
Combo1.AddItem "*.flc"
Combo1.AddItem "*.fli"
Combo1.AddItem "*.avi"
Combo1.AddItem "*.jpg"
Combo1.AddItem "*.bmp"
Combo1.AddItem "*.gif"
Combo1.AddItem "*.txt"
Combo1.AddItem "*.rtf"
'Combo1.AddItem "*.dat"
Combo1.Text = Combo1.List(0)
File1.Pattern = Combo1.Text
Looptext.Text = Str$(HScroll1.Value)
FAText.Text = Str$(HScroll2.Value)
PTText.Text = Str$(HScroll3.Value)
LedForm.Caption = "建立描述文件"
Drive1.Drive = "c:"
If flagofModify = True Then
  LedForm.Caption = "修改描述文件"
  openTempFilename
  TransfertoRecords
End If
If flagofModify = True Then
   flagofModify = False
End If
LedForm.Top = gtop + gheight
LedForm.Left = 2500
Dim i As Integer
For i = 0 To 1000
    Records(i).filename = Space(100)
    Records(i).TXTFontName = Space(30)
Next
End Sub

Private Sub HScroll1_Change()
Looptext.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
FAText.Text = Str$(HScroll2.Value)
End Sub

Private Sub HScroll3_Change()
PTText.Text = Str$(HScroll3.Value)

End Sub

Private Sub List1_Click()
delindex = List1.ListIndex
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     Option1(0).Value = True
     Option1(1).Value = False
Case 1
     Option1(1).Value = True
     Option1(0).Value = False
End Select
End Sub

Private Sub LoopText_Change()
If Val(Looptext.Text) > 10000 Or Val(Looptext.Text) < 1 Then
       action = MsgBox("循环次数范围为(1，10000)!", vbInformation, "消息")
       Looptext.Text = "10000"
End If
HScroll1.Value = Val(Looptext.Text)
End Sub

Private Sub FAText_Change()
If Val(FAText.Text) > 20 Or Val(FAText.Text) < 1 Then
       action = MsgBox("循环次数范围为(1，20)!", vbInformation, "消息")
       FAText.Text = "20"
End If
HScroll2.Value = Val(FAText.Text)
End Sub

Private Sub PTText_Change()
If Val(PTText.Text) > 60 Or Val(PTText.Text) < 1 Then
       action = MsgBox("循环次数范围为(1，60)!", vbInformation, "消息")
       PTText.Text = "60"
End If
HScroll3.Value = Val(PTText.Text)
End Sub
Public Sub TransfertoRecords()
For i% = 0 To Sum - 1 Step 1
     LedForm.List1.AddItem TempArray(i).filename
     Records(i).filename = TempArray(i).filename
     Records(i).NumberofRecord = TempArray(i).NumberofRecord
     Records(i).Loops = TempArray(i).Loops
     Records(i).Speed = TempArray(i).Speed
     Records(i).TXTFontBold = TempArray(i).TXTFontBold
     Records(i).TXTFontItalic = TempArray(i).TXTFontItalic
     Records(i).TXTFontName = TempArray(i).TXTFontName
     Records(i).TXTFontSize = TempArray(i).TXTFontSize
     Records(i).TXTFontUnderline = TempArray(i).TXTFontUnderline
     Records(i).TXTForeColor = TempArray(i).TXTForeColor
     Records(i).TXTMode = TempArray(i).TXTMode
     Records(i).Types = TempArray(i).Types
     SumofRecord = LedForm.List1.ListCount
  Next i
  
End Sub
Public Sub openTempFilename()
Dim LengthofRecord As Integer
Dim temprecord As RecordType
LengthofRecord = Len(temprecord)
FileLength& = FileLen(sTempFilename)
Sum = FileLength / LengthofRecord
ReDim TempArray(Sum) As RecordType
Dim i As Integer
For i = 0 To Sum
    TempArray(i).filename = Space(100)
    TempArray(i).TXTFontName = Space(30)
Next
filenum% = FreeFile
'On Error GoTo dealerror
Open sTempFilename For Random As #filenum Len = LengthofRecord
For i% = 0 To Sum - 1 Step 1
 Get #filenum, i + 1, TempArray(i)
Next i
Close #filenum
End Sub
