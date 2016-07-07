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

'所有的检票口对象
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


'将公历转换成农历，包括星期
Public Function GetChinaDate(pszDate As String) As String

    Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
    Dim curYear, curMonth, curDay, curWeekday
    Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr
    Dim i, m, n, k, isEnd, bit, TheDate

    ''星期名
    WeekName(0) = " * "
    WeekName(1) = "星期日"
    WeekName(2) = "星期一"
    WeekName(3) = "星期二"
    WeekName(4) = "星期三"
    WeekName(5) = "星期四"
    WeekName(6) = "星期五"
    WeekName(7) = "星期六"
    
    ''天干名称
    TianGan(0) = "甲"
    TianGan(1) = "乙"
    TianGan(2) = "丙"
    TianGan(3) = "丁"
    TianGan(4) = "戊"
    TianGan(5) = "己"
    TianGan(6) = "庚"
    TianGan(7) = "辛"
    TianGan(8) = "壬"
    TianGan(9) = "癸"
    
    ''地支名称
    DiZhi(0) = "子"
    DiZhi(1) = "丑"
    DiZhi(2) = "寅"
    DiZhi(3) = "卯"
    DiZhi(4) = "辰"
    DiZhi(5) = "巳"
    DiZhi(6) = "午"
    DiZhi(7) = "未"
    DiZhi(8) = "申"
    DiZhi(9) = "酉"
    DiZhi(10) = "戌"
    DiZhi(11) = "亥"
    
    ''属相名称
    ShuXiang(0) = "鼠"
    ShuXiang(1) = "牛"
    ShuXiang(2) = "虎"
    ShuXiang(3) = "兔"
    ShuXiang(4) = "龙"
    ShuXiang(5) = "蛇"
    ShuXiang(6) = "马"
    ShuXiang(7) = "羊"
    ShuXiang(8) = "猴"
    ShuXiang(9) = "鸡"
    ShuXiang(10) = "狗"
    ShuXiang(11) = "猪"
    
    ''农历日期名
    DayName(0) = "*"
    DayName(1) = "初一"
    DayName(2) = "初二"
    DayName(3) = "初三"
    DayName(4) = "初四"
    DayName(5) = "初五"
    DayName(6) = "初六"
    DayName(7) = "初七"
    DayName(8) = "初八"
    DayName(9) = "初九"
    DayName(10) = "初十"
    DayName(11) = "十一"
    DayName(12) = "十二"
    DayName(13) = "十三"
    DayName(14) = "十四"
    DayName(15) = "十五"
    DayName(16) = "十六"
    DayName(17) = "十七"
    DayName(18) = "十八"
    DayName(19) = "十九"
    DayName(20) = "二十"
    DayName(21) = "廿一"
    DayName(22) = "廿二"
    DayName(23) = "廿三"
    DayName(24) = "廿四"
    DayName(25) = "廿五"
    DayName(26) = "廿六"
    DayName(27) = "廿七"
    DayName(28) = "廿八"
    DayName(29) = "廿九"
    DayName(30) = "三十"
    
    ''农历月份名
    MonName(0) = "*"
    MonName(1) = "正"
    MonName(2) = "二"
    MonName(3) = "三"
    MonName(4) = "四"
    MonName(5) = "五"
    MonName(6) = "六"
    MonName(7) = "七"
    MonName(8) = "八"
    MonName(9) = "九"
    MonName(10) = "十"
    MonName(11) = "十一"
    MonName(12) = "腊"
    
    ''公历每月前面的天数
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
    
    ''农历数据
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
    
    
    ''生成当前公历年、月、日 ==> GongliStr
    curYear = Year(pszDate)
    curMonth = Month(pszDate)
    curDay = Day(pszDate)

    GongliStr = curYear & "年"
    If (curMonth < 10) Then
        GongliStr = GongliStr & "0" & curMonth & "月"
    Else
        GongliStr = GongliStr & curMonth & "月"
    End If
    If (curDay < 10) Then
        GongliStr = GongliStr & "0" & curDay & "日"
    Else
        GongliStr = GongliStr & curDay & "日"
    End If

    ''生成当前公历星期 ==> WeekdayStr
    curWeekday = Weekday(pszDate)
    WeekdayStr = WeekName(curWeekday)

    ''计算到初始时间1921年2月8日的天数：1921-2-8(正月初一)
    TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
    If ((curYear Mod 4) = 0 And curMonth > 2) Then
        TheDate = TheDate + 1
    End If
    
    ''计算农历天干、地支、月、日
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

    ''获取NongliData(m)的第n个二进制位的值
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
    
    ''生成农历天干、地支、属相 ==> NongliStr
    NongliStr = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "年"
    NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"
    
    ''生成农历月、日 ==> NongliDayStr
    If (curMonth < 1) Then
        NongliDayStr = "闰" & MonName(-1 * curMonth)
    Else
        NongliDayStr = MonName(curMonth)
    End If
    NongliDayStr = NongliDayStr & "月"
    
    NongliDayStr = NongliDayStr & DayName(curDay)
    
    '返回农历和星期
    GetChinaDate = MakeDisplayString(NongliStr & NongliDayStr, WeekdayStr)

End Function

Public Function ToBusStatusString(pnBusStatus As Integer) As String
        Select Case pnBusStatus
            Case ST_BusStopped
                ToBusStatusString = "停班"
            Case ST_BusChecking
                ToBusStatusString = "正在检票"
            Case ST_BusExtraChecking
                ToBusStatusString = "正在补检"
            Case ST_BusStopCheck
                ToBusStatusString = "停止检票"
            Case ST_BusReplace
                ToBusStatusString = "顶班"
            Case ST_BusSlitpStop
                ToBusStatusString = "并班"
            Case Else
                ToBusStatusString = "候车"
        End Select
End Function
