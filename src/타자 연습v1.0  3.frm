VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "짧은글"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "타자 연습v1.0  3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4755
   ScaleWidth      =   6315
   StartUpPosition =   2  '화면 가운데
   Begin MCI.MMControl mid 
      Height          =   375
      Left            =   3120
      TabIndex        =   40
      Top             =   3000
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3600
      Top             =   3240
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1920
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "메뉴(&E)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "중단하기(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5160
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "시작하기(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   495
      Left            =   3000
      TabIndex        =   39
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "최대타수:"
      Height          =   495
      Left            =   1800
      TabIndex        =   38
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   960
      TabIndex        =   37
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "타수:"
      Height          =   495
      Left            =   240
      TabIndex        =   36
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   35
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "0 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   34
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label8 
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
      Left            =   2520
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "확인 :"
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
      Left            =   1800
      TabIndex        =   32
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "확인 :"
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
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   30
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   5400
      TabIndex        =   29
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   5160
      TabIndex        =   28
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   4920
      TabIndex        =   27
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   4680
      TabIndex        =   26
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   4440
      TabIndex        =   25
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   4200
      TabIndex        =   24
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   3960
      TabIndex        =   23
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   3720
      TabIndex        =   22
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   3480
      TabIndex        =   21
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   3240
      TabIndex        =   19
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   3000
      TabIndex        =   18
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   2760
      TabIndex        =   17
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   2520
      TabIndex        =   16
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   2040
      TabIndex        =   14
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   1800
      TabIndex        =   13
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label5 
      Caption         =   "맞춘갯수:"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   180
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 문장(100), v, k, c, d
Dim cnt As Byte

Private Sub Command1_Click()
Text1.IMEMode = vbIMEModeHangul '한글

Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Label6.Visible = True
Timer1.Enabled = True
Text1.SetFocus
For i = 0 To k
Label1(i).Caption = "x"
Label1(i).FontSize = Text1.FontSize
Label1(i).Visible = True
Timer2.Enabled = False
Next
Timer3.Enabled = True
cnt = 0
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Label6.Visible = False
Timer1.Enabled = False
For i = 0 To k
Label1(i).Visible = False
Next
Timer2.Enabled = False
Timer3.Enabled = False
cnt = 0
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
문장(1) = "말 한마디로 천냥빛을 갚는다."
문장(2) = "낮말은새가듣고 밤말은쥐가듣는다."
문장(3) = "가는말이 고와야 오는 말도 곱다."
문장(4) = "소 귀에 경 읽기"
문장(5) = "낫보고 기역자도 모른다."
문장(6) = "가까운 이웃이 먼 친척보다 낫다."
문장(7) = "가는 날이 장날이다."
문장(8) = "간이 콩알만하다."
문장(9) = "가재는 게 편"
문장(10) = "가지많은 나무에 바람잘 날이 없다."
문장(11) = "간에 기별도 가지 않는다."
문장(12) = "개밥에 도토리"
문장(13) = "간에 기별도 가지 않는다."
문장(14) = "개천에서 용 난다."
문장(15) = "검은 머리 파 뿌리 되도록"
문장(16) = "공든 탑이 무너지랴"
문장(17) = "까마귀 날자 배 떨어진다."
문장(18) = "꿩 대신 닭"
문장(19) = "꿩 먹고 알 먹는다."
문장(20) = "굳은 땅에 물이 고인다."
문장(21) = "누워서 떡 먹기"
문장(22) = "냉수 먹고 이쑤시기"
문장(23) = "돌다리도 두드려 보고 건너라"
문장(24) = "달걀로 바위 치기"
문장(25) = "들으면 병이요 안들으면 약이다."
문장(26) = "땅 짚고 헤엄치기"
문장(27) = "되로 주고 말로 받는다."
문장(28) = "바늘 가는데 실이 간다."
문장(29) = "소 잃고 외양간 고친다."
문장(30) = "수박 겉 햝기"
문장(31) = "아니 되면 조상 탓"
문장(32) = "발 없는 말이 천리 간다."
문장(33) = "물에 빠진 새앙쥐"
문장(34) = "만리길도 한 걸음으로 시작된다."
문장(35) = "마른 하늘에 날벼락"
문장(36) = "아는 것이 병"
문장(37) = "원수는외나무 다리에서 만난다."
문장(38) = "울며 겨자먹기"
문장(39) = "엎어지면 꼬 닿을 때"
문장(40) = "지렁이도 밟으면 꿈틀한다."
문장(41) = "탕약에 감초 빠질까"
문장(42) = "피는 물보다 진하다."
문장(43) = "제도끼에 제 발등을 찍는다."
문장(44) = "제 꾀에 넘어간다."
문장(45) = "정신 일도 하사 불성"
문장(46) = "점잖은 개가 부뚜막에 오른다."
문장(47) = "입술에 침이나 바르지"
문장(48) = "웃는 낮에 침 뱉으랴"
문장(49) = "우물을 파도 한 우물을 파라"
문장(50) = "웃물이맑아야 아랫물이 맑다"
문장(51) = "흥정은 하고 싸움은 말리랬다"
문장(52) = "한 술 밥에 배부르랴?"
문장(53) = "하룻강아지 범 무서운 줄 모른다"
문장(54) = "하늘이 무너져도 솟아날 구멍이 있다"
문장(55) = "하늘의 별따기"
문장(56) = "핑계 없는 무덤이 없다"
문장(57) = "팥을 콩이라 하여도 곧이 듣는다"
문장(58) = "팔은 안으로 굽는다"
문장(59) = "티끌 모아 태산"
문장(60) = "큰 북에서 큰 소리난다"
문장(61) = "큰 방죽도 개미구멍으로 무너진다"
문장(62) = "칼로 물 베기"
문장(63) = "천리 길도 한 걸음부터"
문장(64) = "집에서 새는 바가지는 들에 가서 샌다"
문장(65) = "지렁이도 밟으면 꿈틀한다"
문장(66) = "쥐구멍을 찾는다"
문장(67) = "쥐구멍에도 볕 들 날이 있다"
문장(68) = "좋은 약은 입에 쓰다"
문장(69) = "종로에서 뺨 맞고 한강 가서 눈 흘긴다"
문장(70) = "제비는 작아도 강남을 간다"
문장(71) = "제 버릇 남 줄까"
문장(72) = "제 논에 물대기"
문장(73) = "작은 고추가 더 맵다"
문장(74) = "자랄 나무는 떡잎부터 알아본다"
문장(75) = "자라에게 놀란 놈이 솥뚜겅 보고 놀란다"
문장(76) = "자는 범 코침주기"
문장(77) = "원님 덕에 나팔 분다"
문장(78) = "아무리 바빠도 바늘 허리 매어 못 쓴다"
문장(79) = "아닌 밤중에 홍두깨"
문장(80) = "십 년이면 강산도 변한다"
문장(81) = "아는 길도 물어 가라"
문장(82) = "아니 땐 굴뚝에 연기날까"
문장(83) = "식은 죽 먹기"
문장(84) = "시작이 반이다"
문장(85) = "시루에 물 퍼붓기"
문장(86) = "숭어가 뛰니까 망둥이도 뛴다"
문장(87) = "수박 겉 햝기"
문장(88) = "송곳도 끝부터 들어간다"
문장(89) = "소 잃고 외양간 고친다"
문장(90) = "우물 안 개구리"
문장(91) = "울며 겨자먹기"
문장(92) = "소경이 개천 나무란다"
문장(93) = "옷이 날개다"
문장(94) = "옥에도 티가 있다"
문장(95) = "업은 아기 삼년 찾는다"
문장(96) = "어물전 망신은 꼴뚜기가 시킨다"
문장(97) = "얕은 내도 깊게 건너라"
문장(98) = "다 된 죽에 코 빠졌다"
문장(99) = "달면 삼키고 쓰면 뱉는다"
문장(100) = "등잔 밑이 어둡다"


a = Int(Rnd(1) * 100)
Label6.Caption = 문장(a)
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Label6.Visible = False
k = Len(문장(a))
v = Text1.Text
For i = 0 To k
Label1(i).Visible = False
Timer3.Enabled = False
Next
Timer2.Enabled = False
Label13.Caption = 0
mid.FileName = App.Path + "\THEME ME.wav"
mid.Command = "open"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Timer2.Enabled = True
Label2.Caption = Text1.Text
If KeyAscii = 13 Then
If Trim(Text1.Text) <> "" Then
If cnt <> 0 Then
Select Case cnt
Case 1 To 2
c = 2
Case 3 To 4
c = 1.9
Case 5 To 6
c = 1.8
Case 6 To 7
c = 1.7
Case 8 To 9
c = 1.6
Case 10 To 11
c = 1.5
Case 12 To 13
c = 1.4
Case 14 To 15
c = 1.3
Case 16 To 17
c = 1.2
Case 18 To 19
c = 1.1
Case 20 To 21
c = 1
End Select
Label10.Caption = Round(((k / cnt) * 60) * c * -1) * -1
If Label12.Caption < Label10.Caption Then
Label12.Caption = Label10.Caption
mid.Command = "prev"
mid.Command = "play"
End If

Label13.Caption = Label13.Caption + 1
        For i = 0 To k
        Label1(i).Caption = ""
        Next
       ' Label10.Caption = Int((d / k) * 0.1)
        If Label6.Caption = Trim(Text1.Text) Then
           Label4.Caption = Label4.Caption + 1
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           End If
           k = Len(문장(a))
           For i = 0 To k
           Label1(i).Caption = "x"
           Next
           Label8.Caption = "맞았다"
        ElseIf v <> Label6.Caption Then
         Label8.Caption = "틀렸다"
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
          End If
           k = Len(문장(a))
           
                      For i = 0 To k
           Label1(i).Caption = "x"
                       
            Beep
           Next
        End If
        End If
    Text1.Text = ""
    Text1.SetFocus
    cnt = 0
End If
End If
End Sub

Private Sub Timer1_Timer()

k = Len(Label6.Caption)

For i = 0 To k
    If Left(Text1.Text, i + 1) = Left(Label6.Caption, i + 1) Then
        Label1(i).Caption = "o"
        End If
         Next
For i = 0 To k
If Left(Text1.Text, i + 1) <> Left(Label6.Caption, i + 1) Then
       Label1(i).Caption = "x"
        End If
        Next
End Sub

Private Sub Timer2_Timer()
'd = Timer2.Interval + Timer2.Interval + 0.1
End Sub

Private Sub Timer3_Timer()
cnt = cnt + 1
End Sub
