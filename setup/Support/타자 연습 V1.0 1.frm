VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "산성비"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "타자 연습 V1.0 1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10005
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "도움말"
      Height          =   495
      Left            =   7800
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "영문 타자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "한글 타자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "점수 보기"
      Height          =   495
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "난이도 선택"
      Height          =   1095
      Left            =   7680
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "초급자"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "중급자"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "상급자"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "메뉴(E)"
      Height          =   495
      Left            =   3960
      Style           =   1  '그래픽
      TabIndex        =   12
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "중단하기(&P)"
      Height          =   495
      Left            =   6000
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "시작하기(&S)"
      Height          =   495
      Left            =   6000
      MaskColor       =   &H00FFC0FF&
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   4320
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   28
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "입력 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   23
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   22
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "이름 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8640
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7320
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   6600
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   9960
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label3 
      Caption         =   "남은 갯수 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "맞은 갯수 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 낱말(200), i, 맞춘갯수(10), 점수(10), 이름(10), 순위(10), 난이도(10), cnt, J, 영문(100), 종류(10)

Private Sub Command1_Click()
Randomize
Label4.Caption = 0
Label5.Caption = 10
For i = 0 To 6
If Command5.Enabled = False Then
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
ElseIf Command6.Enabled = False Then
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Top = 0
End If
Next
Command1.Enabled = False
Command2.Enabled = True
For i = 0 To 6
Label1(i).Visible = True
Next
Timer1.Enabled = True
Text1.SetFocus
Frame1.Visible = False
Frame1.Enabled = False
Text2.Visible = False
Text2.Enabled = False
If Command5.Enabled = True Then
If Command6.Enabled = True Then
MsgBox "타자 종류를 선택해 주세요." + Chr(13) + Chr(13) _
       + "맨 밑에 한글과 영어중 선택해 주세요", vbCritical
Frame1.Visible = True
Frame1.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
Text2.Visible = True
Text2.Enabled = True
End If
End If
Label9.Caption = Text2.Text
End Sub

Private Sub Command2_Click()
Randomize
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
For i = 0 To 6
Label1(i).Visible = False
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)

a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Top = 0
Next
Frame1.Visible = True
Frame1.Enabled = True
Text2.Visible = True
Text2.Enabled = True
Label6.Visible = True
Label6.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Command4_Click()

     ReDim PLAYER(1 To 11) As String
     ReDim SCORE(1 To 11) As Integer
     Dim PLAYERNAME As String
     Dim PLAYERTEMP As String
     Dim SCORETEMP As Integer
     Dim MAINLOOP As Integer
     Dim LOOPCTR As Integer
     Dim FOUNDSW As Integer
     Open App.Path & "\xkwk.txt" For Input As #1
     For LOOPCTR = 1 To 10
          Input #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     For LOOPCTR = 1 To 10
          If Val(Label4.Caption) > SCORE(LOOPCTR) Then FOUNDSW = 1
     Next LOOPCTR
     If FOUNDSW = 1 Then
          PLAYERNAME = InputBox$("이름을 입력하세요")
          If PLAYERNAME = "" Then PLAYERNAME = "이름없음"
          PLAYER(11) = PLAYERNAME
          SCORE(11) = Val(Label4.Caption)
     End If
     For MAINLOOPCTR = 1 To 10
          For LOOPCTR = 1 To 10
               If SCORE(LOOPCTR) < SCORE(LOOPCTR + 1) Then
                    PLAYERTEMP = PLAYER(LOOPCTR)
                    PLAYER(LOOPCTR) = PLAYER(LOOPCTR + 1)
                    PLAYER(LOOPCTR + 1) = PLAYERTEMP
                    SCORETEMP = SCORE(LOOPCTR)
                    SCORE(LOOPCTR) = SCORE(LOOPCTR + 1)
                    SCORE(LOOPCTR + 1) = SCORETEMP
               End If
          Next LOOPCTR
     Next MAINLOOPCTR
     Open "xkwk.txt" For Output As #1
     MESSAGE = "높은 점수" + Chr$(13)
     For LOOPCTR = 1 To 10
          MESSAGE = MESSAGE + Chr$(13)
          If SCORE(LOOPCTR) = SCORECTR Then
               MESSAGE = MESSAGE + "-> " + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          Else
               MESSAGE = MESSAGE + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          End If
          Write #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     MsgBox (MESSAGE)


Timer1.Enabled = False
'cnt = cnt + 1
'맞춘갯수(cnt) = Label4.Caption
'이름(cnt) = Text2.Text
'점수(cnt) = 맞춘갯수(cnt) * 10
'If Command5.Enabled = True Then
'종류(cnt) = "영문"
'ElseIf Command6.Enabled = True Then
'종류(cnt) = "한글"
'End If
'If Option1 = True Then
'난이도(cnt) = "상급자"
'ElseIf Option2 = True Then
'난이도(cnt) = "중급자"
'ElseIf Option3 = True Then
'난이도(cnt) = "초급자"
'Else
'난이도(cnt) = "보통"
'End If
'Form2.Show
'Form2.Print "순위    이름    맞춘갯수     점수    난이도     종류"

'For i = 1 To cnt
'순위(i) = 1
'Next

'For i = 1 To cnt
'For J = 1 To cnt
'If 점수(i) > 점수(J) Then
'    순위(J) = 순위(J) + 1
'End If
'Next
'Next


'For i = 1 To cnt - 1
'For J = i + 1 To cnt
'If 맞춘갯수(i) < 맞춘갯수(J) Then
'   im = 순위(i)
'   순위(i) = 순위(J)
'   순위(J) = im
   
 '  im = 이름(i)
 '  이름(i) = 이름(J)
 '  이름(J) = im
 '
 '  im = 맞춘갯수(i)
 '  맞춘갯수(i) = 맞춘갯수(J)
 '  맞춘갯수(J) = im
 '
 '  im = 점수(i)
 '  점수(i) = 점수(J)
 '  점수(J) = im
 '
 '  im = 난이도(i)
 '  난이도(i) = 난이도(J)
 '  난이도(J) = im
 '
 '  im = 종류(i)
 '  종류(i) = 종류(J)
 '  종류(J) = im
 '  End If
 '  Next
 '  Next

'For i = 1 To cnt
'Form2.Print Tab(2); 순위(i);
'Form2.Print Tab(7); 이름(i);
'Form2.Print Tab(16); 맞춘갯수(i);
'Form2.Print Tab(26); 점수(i);
'Form2.Print Tab(32); 난이도(i);
'Form2.Print Tab(42); 종류(i)
'Next

End Sub

Private Sub Command5_Click()
Randomize
Command1.Enabled = True
Command2.Enabled = False
Text1.IMEMode = vbIMEModeHangul '한글
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


For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Label10.Caption = "1단계"
End Sub

Private Sub Command6_Click()
Randomize
Command1.Enabled = True
Command2.Enabled = False
Text1.IMEMode = vbIMEModeAlpha  '영문
영문(1) = "computer"
영문(2) = "play"
영문(3) = "can"
영문(4) = "mouse"
영문(5) = "key"
영문(6) = "english"
영문(7) = "knight"
영문(8) = "korea"
영문(9) = "man"
영문(10) = "fool"
영문(11) = "say"
영문(12) = "star"
영문(13) = "to"
영문(14) = "enjoy"
영문(15) = "car"
영문(16) = "good"
영문(17) = "baskstball"
영문(18) = "team"
영문(19) = "time"
영문(20) = "join"
영문(21) = "soccer"
영문(22) = "save"
영문(23) = "load"
영문(24) = "prefer"
영문(25) = "love"
영문(26) = "like"
영문(27) = "are"
영문(28) = "do"
영문(29) = "well"
영문(30) = "kill"
영문(31) = "king"
영문(32) = "key"
영문(33) = "short"
영문(34) = "between"
영문(35) = "start"
영문(36) = "shout"
영문(37) = "cross"
영문(38) = "finger"
영문(39) = "wish"
영문(40) = "luck"
영문(41) = "same"
영문(42) = "tie"
영문(43) = "saveral"
영문(44) = "point"
영문(45) = "stand"
영문(46) = "both"
영문(47) = "lose"
영문(48) = "win"
영문(49) = "because"
영문(50) = "get"
영문(51) = "one"
영문(52) = "best"
영문(53) = "country"
영문(54) = "anyway"
영문(55) = "there"
영문(56) = "between"
영문(57) = "go"
영문(58) = "nothing"
영문(59) = "would"
영문(60) = "have"
영문(61) = "cold"
영문(62) = "that"
영문(63) = "shall"
영문(64) = "sounds"
영문(65) = "see"
영문(66) = "you"
영문(67) = "then"
영문(68) = "invite"
영문(69) = "present"
영문(70) = "send"
영문(71) = "doctor"
영문(72) = "glass"
영문(73) = "bed"
영문(74) = "rest"
영문(75) = "roof"
영문(76) = "porch"
영문(77) = "walk"
영문(78) = "into"
영문(79) = "jump"
영문(80) = "full"
영문(81) = "wonderful"
영문(82) = "unhappy"
영문(83) = "happy"
영문(84) = "week"
영문(85) = "much"
영문(86) = "what"
영문(87) = "movi"
영문(88) = "rise"
영문(89) = "body"
영문(90) = "age"
영문(91) = "egg"
영문(92) = "hour"
영문(93) = "also"
영문(94) = "clear"
영문(95) = "healthy"
영문(96) = "look"
영문(97) = "matter"
영문(98) = "mother"
영문(99) = "father"
영문(100) = "easy"

Label10.Caption = "1단계"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
End Sub

Private Sub Command7_Click()
MsgBox "< 이슬비 도움말 >" + Chr(13) + Chr(13) _
       + "이슬비도 제가 자작한 프로그램 입니다." + Chr(13) + Chr(13) _
       + "처음에 시작할때 맨 밑에있는 한글과 영문중에서 선택해 주세요." + Chr(13) + Chr(13) _
       + "난이도 조절도 해주세요. 그리고 프로그램을 끄면 점수가 지워집니다." + Chr(13) + Chr(13) _
       + "점수계산을 포함해 문제가 많군요. 앞으로 계속 버전업할 생각입니다." + Chr(13) + Chr(13) _
       + "맞춘갯수에 따라 11단계로 나누어 집니다." + Chr(13) + Chr(13) _
       + "즐겁게 하세요..^^" + Chr(13) + Chr(13)

End Sub

Private Sub Form_Load()
Command1.Enabled = True
Command2.Enabled = False
Command5.Enabled = True
MsgBox "지구에 강력한 산성을 띄고있는 물이 떨어지고 있다" + Chr(13) + Chr(13) _
        + "우리는 산성비를 막아야 한다."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Randomize
Label7.Caption = Text1.Text
If KeyAscii = 13 Or KeyAscii = 32 Then
For i = 0 To 6

        If Label1(i).Caption = Trim(Text1.Text) Then
           Label4.Caption = Label4.Caption + 1
           Label1(i).Top = 0

       If Command5.Enabled = False Then
      
           a = Int(Rnd(1) * 200)
           Label1(i).Caption = 낱말(a)
           If Label1(i).Caption = "" Then
           a = Int(Rnd(1) * 200)
           Label1(i).Caption = 낱말(a)
           End If
           ElseIf Command6.Enabled = False Then
             a = Int(Rnd(1) * 100)
           Label1(i).Caption = 영문(a)
          
End If
           
            End If
              Next
    Text1.Text = ""
    Text1.SetFocus

If Command5.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2단계"
Label10.Caption = "2단계"
For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3단계"
Label10.Caption = "3단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4단계"
Label10.Caption = "4단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5단계"
Label10.Caption = "5단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6단계"
Label10.Caption = "6단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7단계"
Label10.Caption = "7단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8단계"
Label10.Caption = "8단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9단계"
Label10.Caption = "9단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10단계"
Label10.Caption = "10단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "마지막 단계"
Label10.Caption = "마지막 단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
End If
End If



If Command6.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2단계"
Label10.Caption = "2단계"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3단계"
Label10.Caption = "3단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4단계"
Label10.Caption = "4단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5단계"
Label10.Caption = "5단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6단계"
Label10.Caption = "6단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7단계"
Label10.Caption = "7단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8단계"
Label10.Caption = "8단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9단계"
Label10.Caption = "9단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10단계"
Label10.Caption = "10단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "마지막 단계"
Label10.Caption = "마지막 단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
End If
End If

End If
End Sub

Private Sub Timer1_Timer()
Randomize

Timer1.Enabled = True


For i = 0 To 6
Label1(i).Visible = True
If Option1 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 25)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 30)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 35)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 40)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 40)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 45)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 45)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 45)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 50)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 55)
End If
ElseIf Option2 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 15)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 20)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 25)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 30)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(22) * 32)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 35)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(28) * 38)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 40)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 45)
End If
ElseIf Option3 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(1) * 5)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 10)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 15)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 25)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 32)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 34)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 38)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 38)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 40)
End If
End If
Next

For i = 0 To 6
        
         If Label1(i).Top > 5880 Then
            Label5.Caption = Label5.Caption - 1
            Label1(i).Top = 0
            If Command5.Enabled = False Then
            a = Int(Rnd(1) * 200)
            Label1(i).Caption = 낱말(a)
            ElseIf Command6.Enabled = False Then
            a = Int(Rnd(1) * 100)
            Label1(i).Caption = 영문(a)
            End If
            ElseIf Label5.Caption = 0 Then
            Command1.Enabled = True
            Command2.Enabled = False
            Timer1.Enabled = False
            Label1(i).Visible = False
            MsgBox "Game over"
            Exit Sub
            ElseIf Label5.Caption < 0 Then
            Label5.Caption = 0
            
                                    
End If
        
            
Next

If Command5.Enabled = False Then
If Label1(0).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
End If

If Command6.Enabled = False Then
If Label1(0).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
End If


End Sub

Private Sub Timer2_Timer()


End Sub
