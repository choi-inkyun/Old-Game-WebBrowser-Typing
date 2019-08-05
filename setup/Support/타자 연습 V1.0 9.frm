VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "달리기"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "타자 연습 V1.0 9.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   7365
   ScaleWidth      =   8805
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      Caption         =   "기록판"
      Height          =   1215
      Left            =   2400
      TabIndex        =   27
      Top             =   6000
      Width           =   4215
      Begin VB.Label Label26 
         Alignment       =   2  '가운데 맞춤
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   2  '가운데 맞춤
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   2  '가운데 맞춤
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "99"
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "99"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "99"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "초급자"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "중급자"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "상급자"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   6480
      Top             =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "난이도"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "초급자"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "중급자"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "상급자"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "도움말"
      Height          =   495
      Left            =   7320
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Timer Timer8 
      Interval        =   200
      Left            =   6000
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "처음부터"
      Height          =   495
      Left            =   5640
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "메뉴"
      Height          =   495
      Left            =   7320
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "시작하기(&S)"
      Height          =   495
      Left            =   5640
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   5520
      Top             =   960
   End
   Begin VB.Timer Timer6 
      Interval        =   50
      Left            =   5040
      Top             =   960
   End
   Begin VB.Timer Timer5 
      Interval        =   30
      Left            =   4560
      Top             =   960
   End
   Begin VB.Timer Timer4 
      Interval        =   30
      Left            =   4080
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Interval        =   30
      Left            =   3600
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   3120
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   2640
      Top             =   960
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label10 
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "이름 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label13 
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   7320
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      X1              =   7920
      X2              =   7920
      Y1              =   -240
      Y2              =   4320
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":030A
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":0BD4
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":149E
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":1D68
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":2632
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "타자 연습 V1.0 9.frx":2EFC
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "단어 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 낱말(200), k, b, c, d, e, f, g, h, i
Dim cnt As Byte
Dim aaa1 As Byte
Dim aaa2 As Byte
Dim aaa3 As Byte
Dim nam
Private Sub Command1_Click()
Label9.Caption = "포켓몬 달리기 재경기가 있겠습니다."
Text1.IMEMode = vbIMEModeHangul '한글
Text1.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer9.Enabled = True
Label1.Visible = True
Text1.SetFocus
Text1.Text = ""
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label2.Visible = True
Command1.Enabled = False
Command3.Enabled = True
Label9.Caption = "경기가 시작되었습니다"
Timer8.Enabled = True
Frame1.Visible = False
Label10.Visible = True
Label10.Caption = Text2.Text
nam = Text2.Text
If Text2.Text = "" Then
nam = "피카츄"
Label10 = "피카츄"
End If
Text2.Visible = False

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Frame1.Visible = True
Image1.Left = 120
Image2.Left = 120
Image3.Left = 120
Image4.Left = 120
Image5.Left = 120
Image6.Left = 120
Timer1.Enabled = False
Text1.Enabled = False
Command1.Enabled = True
Command3.Enabled = False
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
cnt = 0
Text2.Visible = True
Label1.Visible = False
End Sub

Private Sub Command4_Click()
MsgBox "<달리기 도움말>" + Chr(13) + Chr(13) _
       + "사용자는 " & nam & "를 조종하게 됩니다." + Chr(13) + Chr(13) _
       + "단어에 써있는 단어를 맞추시면 됩니다." + Chr(13) + Chr(13) _
       + "재미있게 하세요..~.^^.."
End Sub

Private Sub Form_Load()
Label9.Caption = "포켓몬 나라에 달리기 대회가 열였습니다. 선수들 준비."
 낱말(1) = "컴퓨터"
낱말(2) = "대한민국"
낱말(3) = "미나리"
낱말(4) = "타자"
낱말(5) = "워드"
낱말(6) = "간지럽다"
낱말(7) = "멍게"
낱말(8) = "해삼"
낱말(9) = "말미잘"
낱말(10) = "배"
낱말(11) = "일요일"
낱말(12) = "키보드"
낱말(13) = "프린터"
낱말(14) = "스캐너"
낱말(15) = "햇빛"
낱말(16) = "축구"
낱말(17) = "농구"
낱말(18) = "배구"
낱말(19) = "비치"
낱말(20) = "탁구"
낱말(21) = "피구"
낱말(22) = "핸드볼"
낱말(23) = "설악산"
낱말(24) = "금강산"
낱말(25) = "미국"
낱말(26) = "달"
낱말(27) = "토마토"
낱말(28) = "수박"
낱말(29) = "딸기"
낱말(30) = "시계"
낱말(31) = "일어나다"
낱말(32) = "개학"
낱말(33) = "방학"
낱말(34) = "게으름"
낱말(35) = "늦잠"
낱말(36) = "숙제"
낱말(37) = "게임"
낱말(38) = "마우스"
낱말(39) = "스피커"
낱말(40) = "그래픽"
낱말(41) = "궁서"
낱말(42) = "칠판"
낱말(43) = "너털웃음"
낱말(44) = "바람"
낱말(45) = "빛"
낱말(46) = "어둠"
낱말(47) = "무"
낱말(48) = "불"
낱말(49) = "물"
낱말(50) = "땅"
낱말(51) = "노루"
낱말(52) = "노르스름"
낱말(53) = "노구"
낱말(54) = "뼈"
낱말(55) = "걸음"
낱말(56) = "갑작스레"
낱말(57) = "희망"
낱말(58) = "꿈"
낱말(59) = "천상천하"
낱말(60) = "황태자"
낱말(61) = "사고방식"
낱말(62) = "살림"
낱말(63) = "신화"
낱말(64) = "팬"
낱말(65) = "책상"
낱말(66) = "칠판"
낱말(67) = "볼펜"
낱말(68) = "만년필"
낱말(69) = "공책"
낱말(70) = "어머니"
낱말(71) = "아버지"
낱말(72) = "삼촌"
낱말(73) = "고모"
낱말(74) = "이모"
낱말(75) = "할머니"
낱말(76) = "할아버지"
낱말(77) = "도시"
낱말(78) = "버스"
낱말(79) = "지하철"
낱말(80) = "식초"
낱말(81) = "설탕"
낱말(82) = "소금"
낱말(83) = "나트륨"
낱말(84) = "금"
낱말(85) = "식물"
낱말(86) = "동물"
낱말(87) = "고양이"
낱말(88) = "사자"
낱말(89) = "호랑이"
낱말(90) = "하이에나"
낱말(91) = "인터넷"
낱말(92) = "네트워크"
낱말(93) = "서류"
낱말(94) = "휴지통"
낱말(95) = "김치"
낱말(96) = "빵"
낱말(97) = "만화"
낱말(98) = "실로폰"
낱말(99) = "바이올린"
낱말(100) = "떡"
낱말(101) = "안경"
낱말(102) = "쌀"
낱말(103) = "교과서"
낱말(104) = "소설"
낱말(105) = "책"
낱말(106) = "영화"
낱말(107) = "보리"
낱말(108) = "의자"
낱말(109) = "걸상"
낱말(110) = "디스켓"
낱말(111) = "참고서"
낱말(112) = "시험"
낱말(113) = "자격증"
낱말(114) = "아이고"
낱말(115) = "연습"
낱말(116) = "대회"
낱말(117) = "텔레비전"
낱말(118) = "방정맞다"
낱말(119) = "아시아"
낱말(120) = "세계"
낱말(121) = "애국가"
낱말(122) = "하느님"
낱말(123) = "무궁화"
낱말(124) = "산"
낱말(125) = "어지럽다"
낱말(126) = "황당하다"
낱말(127) = "총"
낱말(128) = "살인"
낱말(129) = "죽다"
낱말(130) = "강도"
낱말(131) = "범죄"
낱말(132) = "엄마"
낱말(133) = "아빠"
낱말(134) = "얼다"
낱말(135) = "동상"
낱말(136) = "고드름"
낱말(137) = "눈"
낱말(138) = "비"
낱말(139) = "에메랄드"
낱말(140) = "다이아"
낱말(141) = "카드"
낱말(142) = "봄"
낱말(143) = "여름"
낱말(144) = "가을"
낱말(145) = "겨울"
낱말(146) = "춘하추동"
낱말(147) = "논"
낱말(148) = "밭"
낱말(149) = "추수"
낱말(150) = "아침"
낱말(151) = "점심"
낱말(152) = "저녁"
낱말(153) = "밤"
낱말(154) = "열쇠"
낱말(155) = "구두"
낱말(156) = "신발"
낱말(157) = "운동화"
낱말(158) = "베이직"
낱말(159) = "모임"
낱말(160) = "동아리"
낱말(161) = "동호회"
낱말(162) = "영장"
낱말(163) = "칼"
낱말(164) = "영어"
낱말(165) = "명언"
낱말(166) = "하늘"
낱말(167) = "사랑"
낱말(168) = "기억하다"
낱말(169) = "영혼"
낱말(170) = "과자"
낱말(171) = "사전"
낱말(172) = "불가능"
낱말(173) = "밀레니엄"
낱말(174) = "영재"
낱말(175) = "얘기"
낱말(176) = "오래되다"
낱말(177) = "오도방정"
낱말(178) = "웃다"
낱말(179) = "물"
낱말(180) = "왕궁"
낱말(181) = "귀족"
낱말(182) = "양반"
낱말(183) = "돈"
낱말(184) = "외롭다"
낱말(185) = "외가댁"
낱말(186) = "벌레"
낱말(187) = "우물"
낱말(188) = "우등생"
낱말(189) = "우두머리"
낱말(190) = "우유"
낱말(191) = "운동"
낱말(192) = "원두막"
낱말(193) = "원근감"
낱말(194) = "위기"
낱말(195) = "월식"
낱말(196) = "생명"
낱말(197) = "나무"
낱말(198) = "유언"
낱말(199) = "투자"
낱말(200) = "한결같다"

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Command1.Enabled = True
Command3.Enabled = False
Label1.Visible = False
k = Len(낱말(a))
b = k * 40
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Text1.Enabled = False
Timer8.Enabled = False
Frame1.Visible = True
cnt = 0
Timer9.Enabled = False
Text2.Visible = True
Text2.Text = ""
Label10.Visible = False
Label10.Caption = ""
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
If Trim(Text1.Text) = Label1.Caption Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
End If
k = Len(낱말(a))
b = k * Int(Rnd(140) * 170)
Image6.Left = Image6.Left + b
Else
If Text1.Text <> "" Then
Image6.Left = Image6.Left - Int(Rnd(40) * 70)
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
End If
End If
End If
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Image1.Left = Image1.Left + Int(Rnd(c) * d)
If Image1.Left >= 7560 Then
Label3.Visible = True
Label11.Visible = True
Timer1.Enabled = False
Image1.Left = 8040
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label3.Caption = "1등"
Label9.Caption = "이상해씨가 1등을 하였습니다"
Label11.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "이상해씨"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "이상해씨"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "이상해씨"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If


End If
End If
End If
End If
End If
If Label4.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "이상해씨는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "이상해씨는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "이상해씨는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "이상해씨는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "이상해씨는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "이상해씨는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "이상해씨는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "이상해씨는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "이상해씨는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "이상해씨는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "이상해씨는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "이상해씨는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "이상해씨는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "이상해씨는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "이상해씨는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "이상해씨는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "이상해씨는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "이상해씨는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "이상해씨는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "이상해씨는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "이상해씨는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "이상해씨는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "이상해씨는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "이상해씨는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "이상해씨는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
Timer1.Enabled = False
End If
End Sub

Private Sub Timer10_Timer()
End Sub

Private Sub Timer2_Timer()
Image2.Left = Image2.Left + Int(Rnd(d) * e)
If Image2.Left >= 7560 Then
Label4.Visible = True
Label12.Visible = True
d = 0
Timer2.Enabled = False
Image2.Left = 8040
If Label3.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label4.Caption = "1등"
Label9.Caption = "파이리가 1등을 하였습니다"
Label12.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "파이리"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "파이리"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "파이리"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "파이리는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "파이리는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "파이리는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "파이리는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "파이리는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "파이리는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "파이리는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "파이리는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "파이리는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "파이리는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "파이리는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "파이리는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "파이리는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "파이리는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "파이리는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "파이리는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "파이리는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "파이리는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "파이리는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "파이리는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "파이리는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "파이리는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "파이리는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "파이리는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "파이리는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Image3.Left = Image3.Left + Int(Rnd(e) * f)
If Image3.Left >= 7560 Then
Label5.Visible = True
Label13.Visible = True

Timer3.Enabled = False
e = 0
Image3.Left = 8040
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label5.Caption = "1등"
Label9.Caption = "꼬부기가 1등을 하였습니다"
Label13.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "꼬부기"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "꼬부기"
End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "꼬부기"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "꼬부기는 6등을 하였습니다"
End If
Timer3.Enabled = False
End If

End Sub

Private Sub Timer4_Timer()
Image4.Left = Image4.Left + Int(Rnd(f) * g)
If Image4.Left >= 7560 Then
Label6.Visible = True
Label14.Visible = True
Timer4.Enabled = False
f = 0
Image4.Left = 8040
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label6.Caption = "1등"
Label9.Caption = "잠만보가 1등을 하였습니다"
Label14.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "잠만보"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "잠만보"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "잠만보"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "잠만보는 6등을 하였습니다"
End If
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
Image5.Left = Image5.Left + Int(Rnd(g) * c)
If Image5.Left >= 7560 Then
Label7.Visible = True
Label15.Visible = True
Timer5.Enabled = False
g = 0
Image5.Left = 8040
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label8.Caption = "" Then
Label7.Caption = "1등"
Label9.Caption = "피죤이 1등을 하였습니다"
Label15.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "피죤"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "피죤"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "피죤"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label7.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "피죤은 6등을 하였습니다"
End If
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
If Image6.Left >= 7560 Then
Label8.Visible = True
Label16.Visible = True
Timer6.Enabled = False
b = 0
Label1.Visible = False
Image6.Left = 8040
Text1.Enabled = False
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
Label8.Caption = "1등"
Label9.Caption = "" & nam & "가 1등을 하였습니다"
Label16.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = nam

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = nam

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = nam

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
Timer6.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
If Option1.Value = True Then
c = Int(Rnd(15) * 40)
d = Int(Rnd(15) * 40)
e = Int(Rnd(15) * 40)
f = Int(Rnd(15) * 40)
g = Int(Rnd(15) * 40)
End If
If Option2.Value = True Then
c = Int(Rnd(10) * 30)
d = Int(Rnd(10) * 30)
e = Int(Rnd(10) * 30)
f = Int(Rnd(10) * 30)
g = Int(Rnd(10) * 30)
End If
If Option3.Value = True Then
c = Int(Rnd(5) * 20)
d = Int(Rnd(5) * 20)
e = Int(Rnd(5) * 20)
f = Int(Rnd(5) * 20)
g = Int(Rnd(5) * 20)
End If
aaa1 = Val(Label21.Caption)
aaa2 = Val(Label22.Caption)
aaa3 = Val(Label23.Caption)
End Sub

Private Sub Timer8_Timer()
If Image1.Left > Image2.Left Then
If Image1.Left > Image3.Left Then
If Image1.Left > Image4.Left Then
If Image1.Left > Image5.Left Then
If Image1.Left > Image6.Left Then
Select Case Image1.Left
Case 960 To 7200
Label9.Caption = "현재 이상해씨가 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image2.Left > Image1.Left Then
If Image2.Left > Image3.Left Then
If Image2.Left > Image4.Left Then
If Image2.Left > Image5.Left Then
If Image2.Left > Image6.Left Then
Select Case Image2.Left
Case 960 To 7200
Label9.Caption = "현재 파이리가 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image3.Left > Image1.Left Then
If Image3.Left > Image2.Left Then
If Image3.Left > Image4.Left Then
If Image3.Left > Image5.Left Then
If Image3.Left > Image6.Left Then
Select Case Image3.Left
Case 960 To 7200
Label9.Caption = "현재 꼬부기가 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image4.Left > Image1.Left Then
If Image4.Left > Image2.Left Then
If Image4.Left > Image3.Left Then
If Image4.Left > Image5.Left Then
If Image4.Left > Image6.Left Then
Select Case Image4.Left
Case 960 To 7200
Label9.Caption = "현재 잠만보가 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image5.Left > Image1.Left Then
If Image5.Left > Image2.Left Then
If Image5.Left > Image3.Left Then
If Image5.Left > Image4.Left Then
If Image5.Left > Image6.Left Then
Select Case Image5.Left
Case 960 To 7200
Label9.Caption = "현재 피죤이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image6.Left > Image1.Left Then
If Image6.Left > Image2.Left Then
If Image6.Left > Image3.Left Then
If Image6.Left > Image4.Left Then
If Image6.Left > Image5.Left Then
Select Case Image6.Left
Case 960 To 7200
Label9.Caption = "현재 " & nam & "이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If
If Image6.Left <= -200 Then
MsgBox "당신은 경기장에서 퇴장당했습니다" + Chr(13) + Chr(13) _
        + "그렇기 때문에 경기를 처음부터 시작합니다."
Frame1.Visible = True
Image1.Left = 120
Image2.Left = 120
Image3.Left = 120
Image4.Left = 120
Image5.Left = 120
Image6.Left = 120
Timer1.Enabled = False
Text1.Enabled = False
Command1.Enabled = True
Command3.Enabled = False
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
cnt = 0
End If

End Sub

Private Sub Timer9_Timer()
cnt = cnt + 1
End Sub
