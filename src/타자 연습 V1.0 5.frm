VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "타자검정"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "타자 연습 V1.0 5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6855
   ScaleWidth      =   7560
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   127
      Text            =   "100"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6960
      Top             =   1440
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
      Height          =   390
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   6615
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
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   6615
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
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   6615
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
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "시작하기(&S)"
      Height          =   615
      Left            =   600
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "메뉴(&E)"
      Height          =   615
      Left            =   4920
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   360
   End
   Begin VB.CommandButton command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "맹 점"
      Height          =   495
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "소리에 대한 몽상"
      Height          =   495
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "사평역"
      Height          =   495
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "아버지의 땅"
      Height          =   495
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "그들의 새벽"
      Height          =   495
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label9 
      Caption         =   "시간 :"
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
      Left            =   2520
      TabIndex        =   128
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "초"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   126
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "남은 시간 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   125
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lb 
      BorderStyle     =   1  '단일 고정
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
      Index           =   0
      Left            =   240
      TabIndex        =   124
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lb 
      BorderStyle     =   1  '단일 고정
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
      Index           =   1
      Left            =   240
      TabIndex        =   123
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label lb 
      BorderStyle     =   1  '단일 고정
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
      Index           =   2
      Left            =   240
      TabIndex        =   122
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Label lb 
      BorderStyle     =   1  '단일 고정
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
      Index           =   3
      Left            =   240
      TabIndex        =   121
      Top             =   4080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "긴글 연습 종류 고르기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   120
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   119
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   118
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   117
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   116
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   115
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   114
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   113
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   112
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   111
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   110
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   109
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   108
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   107
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3360
      TabIndex        =   106
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3600
      TabIndex        =   105
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   104
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      TabIndex        =   103
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   4320
      TabIndex        =   102
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4560
      TabIndex        =   101
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4800
      TabIndex        =   100
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   5040
      TabIndex        =   99
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      TabIndex        =   98
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5520
      TabIndex        =   97
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5760
      TabIndex        =   96
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6000
      TabIndex        =   95
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6240
      TabIndex        =   94
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   6480
      TabIndex        =   93
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   92
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   91
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   90
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   89
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   88
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   87
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   86
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   85
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   84
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   83
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   82
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   81
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   80
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3360
      TabIndex        =   79
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3600
      TabIndex        =   78
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   77
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      TabIndex        =   76
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   4320
      TabIndex        =   75
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4560
      TabIndex        =   74
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4800
      TabIndex        =   73
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   5040
      TabIndex        =   72
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      TabIndex        =   71
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5520
      TabIndex        =   70
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5760
      TabIndex        =   69
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6000
      TabIndex        =   68
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6240
      TabIndex        =   67
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   6480
      TabIndex        =   66
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   65
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   64
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   63
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   62
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   61
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   60
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   59
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   58
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   57
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   56
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   55
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   54
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   53
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3360
      TabIndex        =   52
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3600
      TabIndex        =   51
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   50
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      TabIndex        =   49
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   4320
      TabIndex        =   48
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4560
      TabIndex        =   47
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4800
      TabIndex        =   46
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   5040
      TabIndex        =   45
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      TabIndex        =   44
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5520
      TabIndex        =   43
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5760
      TabIndex        =   42
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6000
      TabIndex        =   41
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6240
      TabIndex        =   40
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   6480
      TabIndex        =   39
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   36
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   35
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   34
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   33
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   32
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   31
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   30
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   29
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   28
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   27
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   26
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3360
      TabIndex        =   25
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3600
      TabIndex        =   24
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   23
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4080
      TabIndex        =   22
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   4320
      TabIndex        =   21
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4560
      TabIndex        =   20
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4800
      TabIndex        =   19
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   5040
      TabIndex        =   18
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      TabIndex        =   17
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5520
      TabIndex        =   16
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5760
      TabIndex        =   15
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6000
      TabIndex        =   14
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6240
      TabIndex        =   13
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   6480
      TabIndex        =   12
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   6240
      Width           =   495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 긴글, k, rr, gg

Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False

For i = 0 To 3
lb(i).Visible = True
Text1(i).Visible = True
lb(i).Enabled = True
Text1(i).Enabled = True
Next
Text1(0).SetFocus
For i = 0 To 26
Label2(i).Caption = "x"
Label3(i).Caption = "x"
Label4(i).Caption = "x"
Label5(i).Caption = "x"
Label2(i).Visible = True
Label3(i).Visible = True
Label4(i).Visible = True
Label5(i).Visible = True
Next
Label6.Caption = Val(Text2.Text)
gg = Val(Label6.Caption)
Timer2.Enabled = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
If Text2.Text = "" Then
Label6.Caption = 300
End If
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Text2.Visible = False
Text2.Enabled = False
긴글 = " 서울 시내의 제법 번화한 도로변이 대개 그러하듯이 지금 그가 높은 의자에 앉아있는 곳의 대기 속에도 온갖 소리들이 가득 차 있었다. 물론 그 곳은 길거리인지라 차량들의 엔진 소리와 차바퀴가 바닥에 마찰되는 소리, 소리라기보다 소음에 가까운 것들이 주종을 이루고 있긴 했지만 그래도 약간의 주의만 기울인다면 그 외의 더 많은 종류의 소리도 들을 수 있었다. 말하자면 길거리에 앉아 있어야 하는 직업을 가지고 있는 그는 이제 약간의 경력에 의해 시끄러운 자동차들의 소음 속에 파묻혀 있는, 그래서 간간이 단편적으로 드러나는 행인들의 말소리들, 그들의 몸이 내는 소리들, 하여튼 사람들이 만나고 헤어지고 살아가면서 내는 모든 소리들을 놓치지 않고 들을 수 있었다.  그러나 하루 종일 울리는 높운 데시벨의 소리들은 그의 청각 기관을 멍멍하게 했고 머리 속 더 깊이 파고들어서 가끔 두통을 일으켰다. 그러한 시달림 때문인지 아니면 실제로 어떤 이물질 때문인지 귓속이 근지러움을 느낀 그는 한 손에는 배차 시간표를 들고 오른쪽 새끼손가락으로 역시 오른쪽의 귓속을 후볐다. 그러나 귓속의 기묘한 고통은 그의 손가락 끝이 닿을 수 있는 곳 너머에 있었다. < 끝 >"

For i = 0 To 3
lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Next
Command8.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command7.Visible = False
Command5.Visible = False
Command6.Visible = False
Command4.Visible = False
Command8.Visible = False
Label1.Visible = False
Label1.Enabled = False
Command1.Enabled = True
Timer1.Enabled = False
Command1.Visible = True

Command3.Visible = True
Command3.Enabled = True
Command1.Enabled = True

For i = 0 To 26
Label2(i).Visible = True
Label2(i).Enabled = True
Label3(i).Visible = True
Label3(i).Enabled = True
Label4(i).Visible = True
Label4(i).Enabled = True
Label5(i).Visible = True
Label5(i).Enabled = True
Next
Text2.Visible = False
Text2.Enabled = False
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False

End Sub

Private Sub Command5_Click()
Text2.Visible = False
Text2.Enabled = False
긴글 = " 개를 먹는 행위는 일차적으로 소 돼지 양을 먹는 행위와 구별되고, 다른 한편으로 고양이 비둘기 원숭이를 먹는 행위와도 구별된다. 마찬가지로 개가 사람을 보고 짖어 대는 것도 그리 간단한 문제가 아니다. 그것은 소 돼지 양이 사람들에게 울어대는 것과 구별되고, 또한 고양이 원숭이 비둘기가 우는 것과도 구별된다. 이러한 관계가 미묘하다는 것을 동시에 시사하고 있는 것이다.  이런 생각들이 물 속에서 물방울이 보글보글 솟아오르듯이 하릴없이 그에게 떠올랐던 것은 그가 개라는 동물에 쏟은 관심의 덕분이었다. 사실, 생각해 봐야 별로 유쾌할 바 없는 이 동물에 대해 신경을 쓰게 된 데에는 매우 기구하고, 그럴 만한 사연이 있었다. 그것은 그가 개에 도움을 청해야 했던 채무자의 처지에 있었기 때문이었다. 말하자면 그는 개에 허락도 얻지 않고 그 이름을 도용하기 시작했던 것이다.  이 모두는 우선 그의 성격 탓이었다. 그로 말하자면 특징이라곤 담배를 자주 피운다거나, 원색의 옷을 입은 여자를 극도로 싫어한다든가, 걸을 때 고장난 장난감 인형 같은 몸짓이 조금씩 드러나는 것 정도였다. 그러던 그가 어느 날 개성의 필요성에 눈을 뜬 것이다. < 끝 >"

For i = 0 To 3
lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Next
Command8.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command7.Visible = False
Command5.Visible = False
Command6.Visible = False
Command4.Visible = False
Command8.Visible = False
Label1.Visible = False
Label1.Enabled = False
Command1.Enabled = True
Timer1.Enabled = False
Command1.Visible = True
Command3.Visible = True
Command3.Enabled = True
Command1.Enabled = True

For i = 0 To 26
Label2(i).Visible = True
Label2(i).Enabled = True
Label3(i).Visible = True
Label3(i).Enabled = True
Label4(i).Visible = True
Label4(i).Enabled = True
Label5(i).Visible = True
Label5(i).Enabled = True
Next
Text2.Visible = False
Text2.Enabled = False
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False

End Sub

Private Sub Command6_Click()
Text2.Visible = False
Text2.Enabled = False
긴글 = " 막차는 좀처럼 오지 않았다.  별로 복잡한 내용이랄 것도 없는 장부를 마치 꼼꼼히 확인해 보고 나서야 늙은 역장은 돋보기 안경을 벗어 책상 위에 놓고 일어선다.  벌써 삼십 분이나 지났군.  출입문 위쪽에 붙은 낡은 벽시계가 여덟 시 십 오 분을 가리키고 있다. 하긴 뭐 벌써라는 말을 쓰는 것도 새삼스럽다고 그는 고쳐 생각한다. 이렇게 작은 산골 간이역에서 제 시간에 정확히 도착하는 완행 열차를 보기가 그리 쉬운 일은 아님을 익히 알고 있는 탓이다. 더구나 오늘은 눈까지 내리고 있지 않는가.  역장은 손바닥을 비비며 창가로 다가가더니 유리창 너머로 무심히 시선을 던진다. 건널목 옆 외눈박이 수은등이 껑충하게 서서 홀로 눈을 맞으며 희뿌연 얼굴로 땅바닥을 내려다보고 있다. 송이눈이다. 갓난아이의 주먹만한 눈송이들은 어둠저편에 까맣게 숨어 있다가 느닷없이 수은등의 불빛 속에 뛰어 들어오면서 뚱그렇게 놀란 표정을 채 지우지 못한 채 땅바닥으로 곤두박질치고 있다. 굉장한 눈이다. 바람도 그리 없는데 눈발이 비스듬히 비껴날리고 있다. 늙은 역장은 조금은 근심스런 기색으로 유리창에 얼굴을 바짝 대어 본다.     < 끝 >"

For i = 0 To 3
lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Next
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command8.Enabled = False
Command7.Visible = False
Command5.Visible = False
Command6.Visible = False
Command4.Visible = False
Command8.Visible = False
Label1.Visible = False
Label1.Enabled = False
Command1.Enabled = True
Timer1.Enabled = False
Command1.Visible = True
Command3.Visible = True
Command3.Enabled = True
Command1.Enabled = True
For i = 0 To 26
Label2(i).Visible = True
Label2(i).Enabled = True
Label3(i).Visible = True
Label3(i).Enabled = True
Label4(i).Visible = True
Label4(i).Enabled = True
Label5(i).Visible = True
Label5(i).Enabled = True
Next
Text2.Visible = False
Text2.Enabled = False
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False

End Sub

Private Sub Command7_Click()
Text2.Visible = False
Text2.Enabled = False
긴글 = " 쫓겨가는 한 마리 딱정벌레처럼 트럭은 저만큼 들판 가운데로 난 황토길을 따라 느릿느릿 기어가고 있었다. 고르지 못한 노면에서 바퀴가 투어 오를 때마다 덜컹거리는 쇳소리가 들려 왔고 꽁무니로 부옇게 마른 먼지가 피어 올랐다.  덮개 없는 트럭의 뒷간에 홀로 쭈그려 앉은 채 실려 가고 있는 녀석의 모습이 유난히도 자그맣게 오므라들어 있어 보였다. 뒷간에 적대된 알루미늄 식깡들이 이따금 섬뜩할이만큼 차가운 금속성의 광선을 되쏘곤 했다. 풀잎들이 저마다 윤기를 잃어 가고 있는 들녘과 차츰 잿빛으로 퇴색해 가기 시작하는 야산의 정지된 풍경 속에서 그것은 안간힘을 쓰며 집요하게 꿈틀거리고 있는 단 하나의 운동체였다.  '더럽게 운도 없는 녀석이군 전입해 온 지 보름 만에 초상을 치르다니.'  바지를 까내리고 오줌발을 내갈기며 오 일병이 뇌까렸다. 나는 말없이 마른 풀을 짓씹었다. 바로 조금 전에 우리는 그 트럭에서 내렸었다. 야영지를 출발한 지 얼마 되지 않아 차가 마을로 통하는 샛길 입구에 다다랐을 때 선임 탑승자는 차를 세워 우리 둘을 내려 주었던 것이다.  이제 트럭은 들판을 지나 산모퉁이를 마악 꺾어 돌아가려는 참이었다. < 끝 >"

 
For i = 0 To 3
lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Visible = True
Next
Command8.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command7.Visible = False
Command5.Visible = False
Command6.Visible = False
Command4.Visible = False
Command8.Visible = False
Label1.Visible = False
Label1.Enabled = False
Command1.Enabled = True

Timer1.Enabled = False
Command1.Visible = True

Command3.Visible = True
Command3.Enabled = True
Command1.Enabled = True
For i = 0 To 26
Label2(i).Visible = True
Label2(i).Enabled = True
Label3(i).Visible = True
Label3(i).Enabled = True
Label4(i).Visible = True
Label4(i).Enabled = True
Label5(i).Visible = True
Label5(i).Enabled = True
Next
Text2.Visible = False
Text2.Enabled = False
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False

End Sub

Private Sub Command8_Click()
Text2.Visible = False
Text2.Enabled = False
긴글 = " 약기운이 차츰 소진해 가는 마취 상태에서처럼 몽롱한 의식을 후두둑 털어 내며 그녀는 눈을 떴다.  희고 검은 빛깔의 물고기 형상을 하고 균일한 분포로 판박이된 천정의 사방 연속 무늬가 어슴푸레 공중에 걸려 있는 게 맨 먼저 시야에 들어왔다. 현관 바깥에 매달린 외등에서 가느다란 불빛이 유리창으로 새어 들어와 맞은편 벽면으로 날이 잘 다듬어진 비수처럼 음험한 그림자를 드리우고 있었다. 그녀는 메말라 껄끄러운 눈꺼풀을 몇 번인가 깜박거리며 눈의 초점을 맞추려 애를 썼다.  뚜걱, 뚜걱, 뚜거덕.  불현듯 온몸의 털구멍이 한꺼번에 바짝 아가리를 닫고 수축되어 버리는 듯한 긴장감. 그녀는 전신이 풀먹인 무명베처럼 빳빳하게 굳어 가는 느낌이 들었다. 발 소리는 역시 이츰에서 들려 오고 있었다. 두 뼘도 채 못 되는 청정의 콘크리트 두께를 뚫고 발소리는 분명히 그녀의 귀에까지 전달되고 있었다.  뚜걱, 뚜거덕, 뚜걱.  구두 밑창의 두꺼운 뒷축이 시멘트 바닥에 맞부딪쳐서 내는 둔중한 마찰음. 군화를 신고 있거나 굽 높은 투박한 등산화를 신었는지도 모른다. 발소리는 이 날따라 유난히 크도 대담하게 울리는 것 같았다.  아니, 그건 언제나 그랬다.  < 끝 >  "


For i = 0 To 3
lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Next
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command8.Enabled = False
Command7.Visible = False
Command5.Visible = False
Command6.Visible = False
Command4.Visible = False
Command8.Visible = False
Label1.Visible = False
Label1.Enabled = False
Timer1.Enabled = False
Command1.Visible = True
Command3.Visible = True
Command3.Enabled = True
Command1.Enabled = True
For i = 0 To 26
Label2(i).Visible = True
Label2(i).Enabled = True
Label3(i).Visible = True
Label3(i).Enabled = True
Label4(i).Visible = True
Label4(i).Enabled = True
Label5(i).Visible = True
Label5(i).Enabled = True
Next
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False
End Sub


Private Sub Form_Load()

Command1.Enabled = True

Timer1.Enabled = False
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = False
lb(i).Enabled = False
Text1(i).Enabled = False
Next
For i = 0 To 26
Label2(i).Visible = False
Label2(i).Enabled = False
Label3(i).Visible = False
Label3(i).Enabled = False
Label4(i).Visible = False
Label4(i).Enabled = False
Label5(i).Visible = False
Label5(i).Enabled = False
Next
Command1.Visible = False

Command3.Visible = False
Command3.Enabled = False
Command1.Enabled = False
k = 1
Timer2.Enabled = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If rr = -1 Then
    If KeyAscii = 13 Then
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    For i = 0 To 3
        lb(i).Caption = mid(긴글, k, 27)
        k = k + 27
    Next
    End If
End If


If KeyAscii = 13 Then
    
    rr = rr + 1
    Text1(rr).SetFocus
    If rr = 3 Then
    rr = -1
'       If KeyAscii = 13 Then
 '           MsgBox "11"
  '      End If
     End If
End If

End Sub

Private Sub Timer1_Timer()
For i = 0 To 26
If Left(Text1(0).Text, i + 1) = Left(lb(0).Caption, i + 1) Then
        Label2(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(0).Text, i + 1) <> Left(lb(0).Caption, i + 1) Then
       Label2(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(1).Text, i + 1) = Left(lb(1).Caption, i + 1) Then
        Label3(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(1).Text, i + 1) <> Left(lb(1).Caption, i + 1) Then
       Label3(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(2).Text, i + 1) = Left(lb(2).Caption, i + 1) Then
        Label4(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(2).Text, i + 1) <> Left(lb(2).Caption, i + 1) Then
       Label4(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(3).Text, i + 1) = Left(lb(3).Caption, i + 1) Then
        Label5(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(3).Text, i + 1) <> Left(lb(3).Caption, i + 1) Then
       Label5(i).Caption = "x"
        End If
        Next
End Sub

Private Sub Timer2_Timer()
gg = gg - 1
Label6.Caption = gg
If Label6.Caption = 0 Then
MsgBox "시간이 다 되었습니다."
Text1(0).SetFocus
Timer1.Enabled = False
Command1.Enabled = True
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Text1(i).Text = ""
Next
For i = 0 To 26
Label2(i).Visible = False
Label3(i).Visible = False
Label4(i).Visible = False
Label5(i).Visible = False
Next
Timer2.Enabled = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
End If

End Sub
