VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Ÿ�ڰ���"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "Ÿ�� ���� V1.0 5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6855
   ScaleWidth      =   7560
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�����ϱ�(&S)"
      Height          =   615
      Left            =   600
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�޴�(&E)"
      Height          =   615
      Left            =   4920
      Style           =   1  '�׷���
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
      Caption         =   "�� ��"
      Height          =   495
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�Ҹ��� ���� ����"
      Height          =   495
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "����"
      Height          =   495
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�ƹ����� ��"
      Height          =   495
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�׵��� ����"
      Height          =   495
      Left            =   1800
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label9 
      Caption         =   "�ð� :"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���� �ð� :"
      BeginProperty Font 
         Name            =   "����"
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
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
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
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
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
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
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
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      Caption         =   "��� ���� ���� ����"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
Dim ���, k, rr, gg

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
��� = " ���� �ó��� ���� ��ȭ�� ���κ��� �밳 �׷��ϵ��� ���� �װ� ���� ���ڿ� �ɾ��ִ� ���� ��� �ӿ��� �°� �Ҹ����� ���� �� �־���. ���� �� ���� ��Ÿ������� �������� ���� �Ҹ��� �������� �ٴڿ� �����Ǵ� �Ҹ�, �Ҹ���⺸�� ������ ����� �͵��� ������ �̷�� �ֱ� ������ �׷��� �ణ�� ���Ǹ� ����δٸ� �� ���� �� ���� ������ �Ҹ��� ���� �� �־���. �����ڸ� ��Ÿ��� �ɾ� �־�� �ϴ� ������ ������ �ִ� �״� ���� �ణ�� ��¿� ���� �ò����� �ڵ������� ���� �ӿ� �Ĺ��� �ִ�, �׷��� ������ ���������� �巯���� ���ε��� ���Ҹ���, �׵��� ���� ���� �Ҹ���, �Ͽ�ư ������� ������ ������� ��ư��鼭 ���� ��� �Ҹ����� ��ġ�� �ʰ� ���� �� �־���.  �׷��� �Ϸ� ���� �︮�� ���� ���ú��� �Ҹ����� ���� û�� ����� �۸��ϰ� �߰� �Ӹ� �� �� ���� �İ�� ���� ������ �����״�. �׷��� �ô޸� �������� �ƴϸ� ������ � �̹��� �������� �Ӽ��� ���������� ���� �״� �� �տ��� ���� �ð�ǥ�� ��� ������ �����հ������� ���� �������� �Ӽ��� �ĺ���. �׷��� �Ӽ��� �⹦�� ������ ���� �հ��� ���� ���� �� �ִ� �� �ʸӿ� �־���. < �� >"

For i = 0 To 3
lb(i).Caption = mid(���, k, 27)
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
��� = " ���� �Դ� ������ ���������� �� ���� ���� �Դ� ������ �����ǰ�, �ٸ� �������� ����� ��ѱ� �����̸� �Դ� �����͵� �����ȴ�. ���������� ���� ����� ���� ¢�� ��� �͵� �׸� ������ ������ �ƴϴ�. �װ��� �� ���� ���� ����鿡�� ����� �Ͱ� �����ǰ�, ���� ����� ������ ��ѱⰡ ��� �Ͱ��� �����ȴ�. �̷��� ���谡 �̹��ϴٴ� ���� ���ÿ� �û��ϰ� �ִ� ���̴�.  �̷� �������� �� �ӿ��� ������� ���ۺ��� �ھƿ������� �ϸ����� �׿��� ���ö��� ���� �װ� ����� ������ ���� ������ �����̾���. ���, ������ ���� ���� ������ �� ���� �� ������ ���� �Ű��� ���� �� ������ �ſ� �ⱸ�ϰ�, �׷� ���� �翬�� �־���. �װ��� �װ� ���� ������ û�ؾ� �ߴ� ä������ ó���� �־��� �����̾���. �����ڸ� �״� ���� ����� ���� �ʰ� �� �̸��� �����ϱ� �����ߴ� ���̴�.  �� ��δ� �켱 ���� ���� ſ�̾���. �׷� �����ڸ� Ư¡�̶�� ��踦 ���� �ǿ�ٰų�, ������ ���� ���� ���ڸ� �ص��� �Ⱦ��Ѵٵ簡, ���� �� ���峭 �峭�� ���� ���� ������ ���ݾ� �巯���� �� ��������. �׷��� �װ� ��� �� ������ �ʿ伺�� ���� �� ���̴�. < �� >"

For i = 0 To 3
lb(i).Caption = mid(���, k, 27)
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
��� = " ������ ��ó�� ���� �ʾҴ�.  ���� ������ �����̶� �͵� ���� ��θ� ��ġ �Ĳ��� Ȯ���� ���� ������ ���� ������ ������ �Ȱ��� ���� å�� ���� ���� �Ͼ��.  ���� ��� ���̳� ������.  ���Թ� ���ʿ� ���� ���� ���ð谡 ���� �� �� �� ���� ����Ű�� �ִ�. �ϱ� �� ������ ���� ���� �͵� ���ｺ���ٰ� �״� ���� �����Ѵ�. �̷��� ���� ��� ���̿����� �� �ð��� ��Ȯ�� �����ϴ� ���� ������ ���Ⱑ �׸� ���� ���� �ƴ��� ���� �˰� �ִ� ſ�̴�. ������ ������ ������ ������ ���� �ʴ°�.  ������ �չٴ��� ���� â���� �ٰ������� ����â �ʸӷ� ������ �ü��� ������. �ǳθ� �� �ܴ����� �������� �����ϰ� ���� Ȧ�� ���� ������ ��ѿ� �󱼷� ���ٴ��� �����ٺ��� �ִ�. ���̴��̴�. ���������� �ָԸ��� �����̵��� ������� ��İ� ���� �ִٰ� ������� �������� �Һ� �ӿ� �پ� �����鼭 �ױ׷��� ��� ǥ���� ä ������ ���� ä ���ٴ����� ��ι���ġ�� �ִ�. ������ ���̴�. �ٶ��� �׸� ���µ� ������ �񽺵��� �񲸳����� �ִ�. ���� ������ ������ �ٽɽ��� ������� ����â�� ���� ��¦ ��� ����.     < �� >"

For i = 0 To 3
lb(i).Caption = mid(���, k, 27)
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
��� = " �Ѱܰ��� �� ���� ��������ó�� Ʈ���� ����ŭ ���� ����� �� Ȳ����� ���� �������� ���� �־���. ���� ���� ��鿡�� ������ ���� ���� ������ ���ȰŸ��� ��Ҹ��� ��� �԰� �ǹ��Ϸ� �ο��� ���� ������ �Ǿ� �ö���.  ���� ���� Ʈ���� �ް��� Ȧ�� �ޱ׷� ���� ä �Ƿ� ���� �ִ� �༮�� ����� �������� �ڱ׸İ� ���Ƕ��� �־� ������. �ް��� ����� �˷�̴� �ı����� �̵��� �������̸�ŭ ������ �ݼӼ��� ������ �ǽ�� �ߴ�. Ǯ�ٵ��� ������ ���⸦ �Ҿ� ���� �ִ� ���� ���� ������� ����� ���� �����ϴ� �߻��� ������ ǳ�� �ӿ��� �װ��� �Ȱ����� ���� �����ϰ� ��Ʋ�Ÿ��� �ִ� �� �ϳ��� �ü����.  '������ � ���� �༮�̱� ������ �� �� ���� ���� �ʻ��� ġ���ٴ�.'  ������ ����� ���ܹ��� ������� �� �Ϻ��� ����ȴ�. ���� ������ ���� Ǯ�� ���þ���. �ٷ� ���� ���� �츮�� �� Ʈ������ ���Ⱦ���. �߿����� ����� �� �� ���� �ʾ� ���� ������ ���ϴ� ���� �Ա��� �ٴٶ��� �� ���� ž���ڴ� ���� ���� �츮 ���� ���� �־��� ���̴�.  ���� Ʈ���� ������ ���� ������̸� ���� ���� ���ư����� ���̾���. < �� >"

 
For i = 0 To 3
lb(i).Caption = mid(���, k, 27)
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
��� = " ������ ���� ������ ���� ���� ���¿���ó�� ������ �ǽ��� �ĵε� �о� ���� �׳�� ���� ����.  ��� ���� ������ ����� ������ �ϰ� ������ ������ �ǹ��̵� õ���� ��� ���� ���̰� �Ǫ�� ���߿� �ɷ� �ִ� �� �� ���� �þ߿� ���Դ�. ���� �ٱ��� �Ŵ޸� �ܵ�� �����ٶ� �Һ��� ����â���� ���� ���� ������ �������� ���� �� �ٵ���� ���ó�� ������ �׸��ڸ� �帮��� �־���. �׳�� �޸��� �������� ����Ǯ�� �� ���ΰ� ���ڰŸ��� ���� ������ ���߷� �ָ� ���.  �Ѱ�, �Ѱ�, �ѰŴ�.  ������ �¸��� �б����� �Ѳ����� ��¦ �ư����� �ݰ� ����Ǿ� ������ ���� ���尨. �׳�� ������ Ǯ���� ����ó�� �����ϰ� ���� ���� ������ �����. �� �Ҹ��� ���� �������� ��� ���� �־���. �� �µ� ä �� �Ǵ� û���� ��ũ��Ʈ �β��� �հ� �߼Ҹ��� �и��� �׳��� �Ϳ����� ���޵ǰ� �־���.  �Ѱ�, �ѰŴ�, �Ѱ�.  ���� ��â�� �β��� ������ �ø�Ʈ �ٴڿ� �ºε��ļ� ���� ������ ������. ��ȭ�� �Ű� �ְų� �� ���� ������ ���ȭ�� �ž������� �𸥴�. �߼Ҹ��� �� ������ ������ ũ�� ����ϰ� �︮�� �� ���Ҵ�.  �ƴ�, �װ� ������ �׷���.  < �� >  "


For i = 0 To 3
lb(i).Caption = mid(���, k, 27)
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
        lb(i).Caption = mid(���, k, 27)
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
MsgBox "�ð��� �� �Ǿ����ϴ�."
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
