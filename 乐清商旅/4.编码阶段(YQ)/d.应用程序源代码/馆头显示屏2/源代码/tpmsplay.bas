Attribute VB_Name = "Module1"
Option Explicit


Global flagofModify As Boolean
Global CurrentPath As String
Global FlagofAuto As Boolean
Global SavedLEDFile As String
Global TXTfw As Long
Global TXTfh As Long
Global hbr As Long
Global TXTfilenum As InputModeConstants
Global RTFStartPosition As Long
Global RTFStep As Long
Global TwipspercentPixelX As Integer
Global TwipspercentPixelY As Integer
Global avilength As Long
Global Flag As Boolean
Global oldNumberofCurrentFile As Integer
Global NumberofCurrentFile As Integer
Global gleft As Single
Global gtop As Single
Global gwidth As Single
Global gheight As Single
Global currentloops As Integer
Global FirstFlags As Boolean
Global SumofRecords As Integer
Global RecordArray() As RecordType
Global sTempFilename As String
Global iScrollRowofRealTime As Integer
Global iIntervaltimeofRealTime As Integer
Global flagofdatabaseform As Boolean
Public oUser As New ActiveUser
Public oSystem  As New SystemParam
Public oTo As New RTConnection
Public Const szTime = "2000-04-05 08:00:13"
'Public oSellTicketClient  As New SellTicketClient
'Public oGate As Object

'���еļ�Ʊ�ڶ���
'Public oEntrances As Collection

Declare Function ShowFrame Lib "3ds.dll" (ByVal hwnd As Long, ByVal hdc As Long, ByVal filename As String, ByVal x0 As Integer, ByVal y0 As Integer, ByVal width As Integer, ByVal height As Integer, ByVal framenumber As Integer) As Integer
Declare Function GetFrameNum Lib "3ds.dll" (ByVal filename As String, ByVal hwnd As Long) As Integer
Declare Function Flcstop Lib "3ds.dll" (ByVal hwnd As Long) As Integer

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const wFlags = &H2 Or &H1 Or &H40 Or &H10

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Type RecordType
    NumberofRecord As Integer
    filename As String * 100
    Loops As Integer
    Speed As Integer
    Types As Integer
    TXTMode As Integer
    TXTForeColor As Long
    TXTFontName As String * 30
    TXTFontSize As Single
    TXTFontBold As Boolean
    TXTFontItalic As Boolean
    TXTFontUnderline As Boolean
End Type

Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long



Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long


Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long

Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long



Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Type Size
        cx As Long
        cy As Long
End Type


Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long


'������ת����ũ������������
Public Function GetChinaDate(pszDate As String) As String

    Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
    Dim curYear, curMonth, curDay, curWeekday
    Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr
    Dim i, m, n, k, isEnd, bit, TheDate

    ''������
    WeekName(0) = " * "
    WeekName(1) = "������"
    WeekName(2) = "����һ"
    WeekName(3) = "���ڶ�"
    WeekName(4) = "������"
    WeekName(5) = "������"
    WeekName(6) = "������"
    WeekName(7) = "������"
    
    ''�������
    TianGan(0) = "��"
    TianGan(1) = "��"
    TianGan(2) = "��"
    TianGan(3) = "��"
    TianGan(4) = "��"
    TianGan(5) = "��"
    TianGan(6) = "��"
    TianGan(7) = "��"
    TianGan(8) = "��"
    TianGan(9) = "��"
    
    ''��֧����
    DiZhi(0) = "��"
    DiZhi(1) = "��"
    DiZhi(2) = "��"
    DiZhi(3) = "î"
    DiZhi(4) = "��"
    DiZhi(5) = "��"
    DiZhi(6) = "��"
    DiZhi(7) = "δ"
    DiZhi(8) = "��"
    DiZhi(9) = "��"
    DiZhi(10) = "��"
    DiZhi(11) = "��"
    
    ''��������
    ShuXiang(0) = "��"
    ShuXiang(1) = "ţ"
    ShuXiang(2) = "��"
    ShuXiang(3) = "��"
    ShuXiang(4) = "��"
    ShuXiang(5) = "��"
    ShuXiang(6) = "��"
    ShuXiang(7) = "��"
    ShuXiang(8) = "��"
    ShuXiang(9) = "��"
    ShuXiang(10) = "��"
    ShuXiang(11) = "��"
    
    ''ũ��������
    DayName(0) = "*"
    DayName(1) = "��һ"
    DayName(2) = "����"
    DayName(3) = "����"
    DayName(4) = "����"
    DayName(5) = "����"
    DayName(6) = "����"
    DayName(7) = "����"
    DayName(8) = "����"
    DayName(9) = "����"
    DayName(10) = "��ʮ"
    DayName(11) = "ʮһ"
    DayName(12) = "ʮ��"
    DayName(13) = "ʮ��"
    DayName(14) = "ʮ��"
    DayName(15) = "ʮ��"
    DayName(16) = "ʮ��"
    DayName(17) = "ʮ��"
    DayName(18) = "ʮ��"
    DayName(19) = "ʮ��"
    DayName(20) = "��ʮ"
    DayName(21) = "إһ"
    DayName(22) = "إ��"
    DayName(23) = "إ��"
    DayName(24) = "إ��"
    DayName(25) = "إ��"
    DayName(26) = "إ��"
    DayName(27) = "إ��"
    DayName(28) = "إ��"
    DayName(29) = "إ��"
    DayName(30) = "��ʮ"
    
    ''ũ���·���
    MonName(0) = "*"
    MonName(1) = "��"
    MonName(2) = "��"
    MonName(3) = "��"
    MonName(4) = "��"
    MonName(5) = "��"
    MonName(6) = "��"
    MonName(7) = "��"
    MonName(8) = "��"
    MonName(9) = "��"
    MonName(10) = "ʮ"
    MonName(11) = "ʮһ"
    MonName(12) = "��"
    
    ''����ÿ��ǰ�������
    MonthAdd(0) = 0
    MonthAdd(1) = 31
    MonthAdd(2) = 59
    MonthAdd(3) = 90
    MonthAdd(4) = 120
    MonthAdd(5) = 151
    MonthAdd(6) = 181
    MonthAdd(7) = 212
    MonthAdd(8) = 243
    MonthAdd(9) = 273
    MonthAdd(10) = 304
    MonthAdd(11) = 334
    
    ''ũ������
    NongliData(0) = 2635
    NongliData(1) = 333387
    NongliData(2) = 1701
    NongliData(3) = 1748
    NongliData(4) = 267701
    NongliData(5) = 694
    NongliData(6) = 2391
    NongliData(7) = 133423
    NongliData(8) = 1175
    NongliData(9) = 396438
    NongliData(10) = 3402
    NongliData(11) = 3749
    NongliData(12) = 331177
    NongliData(13) = 1453
    NongliData(14) = 694
    NongliData(15) = 201326
    NongliData(16) = 2350
    NongliData(17) = 465197
    NongliData(18) = 3221
    NongliData(19) = 3402
    NongliData(20) = 400202
    NongliData(21) = 2901
    NongliData(22) = 1386
    NongliData(23) = 267611
    NongliData(24) = 605
    NongliData(25) = 2349
    NongliData(26) = 137515
    NongliData(27) = 2709
    NongliData(28) = 464533
    NongliData(29) = 1738
    NongliData(30) = 2901
    NongliData(31) = 330421
    NongliData(32) = 1242
    NongliData(33) = 2651
    NongliData(34) = 199255
    NongliData(35) = 1323
    NongliData(36) = 529706
    NongliData(37) = 3733
    NongliData(38) = 1706
    NongliData(39) = 398762
    NongliData(40) = 2741
    NongliData(41) = 1206
    NongliData(42) = 267438
    NongliData(43) = 2647
    NongliData(44) = 1318
    NongliData(45) = 204070
    NongliData(46) = 3477
    NongliData(47) = 461653
    NongliData(48) = 1386
    NongliData(49) = 2413
    NongliData(50) = 330077
    NongliData(51) = 1197
    NongliData(52) = 2637
    NongliData(53) = 268877
    NongliData(54) = 3365
    NongliData(55) = 531109
    NongliData(56) = 2900
    NongliData(57) = 2922
    NongliData(58) = 398042
    NongliData(59) = 2395
    NongliData(60) = 1179
    NongliData(61) = 267415
    NongliData(62) = 2635
    NongliData(63) = 661067
    NongliData(64) = 1701
    NongliData(65) = 1748
    NongliData(66) = 398772
    NongliData(67) = 2742
    NongliData(68) = 2391
    NongliData(69) = 330031
    NongliData(70) = 1175
    NongliData(71) = 1611
    NongliData(72) = 200010
    NongliData(73) = 3749
    NongliData(74) = 527717
    NongliData(75) = 1452
    NongliData(76) = 2742
    NongliData(77) = 332397
    NongliData(78) = 2350
    NongliData(79) = 3222
    NongliData(80) = 268949
    NongliData(81) = 3402
    NongliData(82) = 3493
    NongliData(83) = 133973
    NongliData(84) = 1386
    NongliData(85) = 464219
    NongliData(86) = 605
    NongliData(87) = 2349
    NongliData(88) = 334123
    NongliData(89) = 2709
    NongliData(90) = 2890
    NongliData(91) = 267946
    NongliData(92) = 2773
    NongliData(93) = 592565
    NongliData(94) = 1210
    NongliData(95) = 2651
    NongliData(96) = 395863
    NongliData(97) = 1323
    NongliData(98) = 2707
    NongliData(99) = 265877
    
    
    ''���ɵ�ǰ�����ꡢ�¡��� ==> GongliStr
    curYear = Year(pszDate)
    curMonth = Month(pszDate)
    curDay = Day(pszDate)

    GongliStr = curYear & "��"
    If (curMonth < 10) Then
        GongliStr = GongliStr & "0" & curMonth & "��"
    Else
        GongliStr = GongliStr & curMonth & "��"
    End If
    If (curDay < 10) Then
        GongliStr = GongliStr & "0" & curDay & "��"
    Else
        GongliStr = GongliStr & curDay & "��"
    End If

    ''���ɵ�ǰ�������� ==> WeekdayStr
    curWeekday = Weekday(pszDate)
    WeekdayStr = WeekName(curWeekday)

    ''���㵽��ʼʱ��1921��2��8�յ�������1921-2-8(���³�һ)
    TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
    If ((curYear Mod 4) = 0 And curMonth > 2) Then
        TheDate = TheDate + 1
    End If
    
    ''����ũ����ɡ���֧���¡���
    isEnd = 0
    m = 0
    
    Do
    If (NongliData(m) < 4095) Then
        k = 11
    Else
        k = 12
    End If
    
    n = k
    Do
    If (n < 0) Then
    Exit Do
    End If

    ''��ȡNongliData(m)�ĵ�n��������λ��ֵ
    bit = NongliData(m)
    For i = 1 To n Step 1
        bit = Int(bit / 2)
    Next
    bit = bit Mod 2
    
    If (TheDate <= 29 + bit) Then
        isEnd = 1
        Exit Do
    End If
    
    TheDate = TheDate - 29 - bit
    
    n = n - 1
    Loop
    
    If (isEnd = 1) Then
        Exit Do
    End If
    
    m = m + 1
    Loop

    curYear = 1921 + m
    curMonth = k - n + 1
    curDay = TheDate
    
    If (k = 12) Then
        If (curMonth = (Int(NongliData(m) / 65536) + 1)) Then
            curMonth = 1 - curMonth
        ElseIf (curMonth > (Int(NongliData(m) / 65536) + 1)) Then
            curMonth = curMonth - 1
        End If
    End If
    
    ''����ũ����ɡ���֧������ ==> NongliStr
    NongliStr = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "��"
    NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"
    
    ''����ũ���¡��� ==> NongliDayStr
    If (curMonth < 1) Then
        NongliDayStr = "��" & MonName(-1 * curMonth)
    Else
        NongliDayStr = MonName(curMonth)
    End If
    NongliDayStr = NongliDayStr & "��"
    
    NongliDayStr = NongliDayStr & DayName(curDay)
    
    '����ũ��������
    GetChinaDate = MakeDisplayString(NongliStr & NongliDayStr, WeekdayStr)

End Function

Public Function ToBusStatusString(pnBusStatus As Integer) As String
        Select Case pnBusStatus
            Case ST_BusStopped
                ToBusStatusString = "ͣ��"
            Case ST_BusChecking
                ToBusStatusString = "���ڼ�Ʊ"
            Case ST_BusExtraChecking
                ToBusStatusString = "���ڲ���"
            Case ST_BusStopCheck
                ToBusStatusString = "ֹͣ��Ʊ"
            Case ST_BusReplace
                ToBusStatusString = "����"
            Case ST_BusSlitpStop
                ToBusStatusString = "����"
            Case Else
                ToBusStatusString = "��"
        End Select
End Function
