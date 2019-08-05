VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "도움말"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "게임4.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6690
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "도움말 종료"
      Height          =   780
      Left            =   6000
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   60
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "게임4.frx":030A
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "게임4.frx":2AAC
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   60
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7560
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   60
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "안녕하세요. Air Force 의 도움말 입니다." + Chr(13) + Chr(13) _
        + "Space = 미사일, Ctrl = 필살기, 각방향키로 움직이시면 됩니다" + Chr(13) + Chr(13) _
        + "1, 2번째 스테이지로 구분되어 있습니다...그럼 재미있게 즐기시길 바랍니다." + Chr(13) + Chr(13) _
        + "버그나 기타 의문사항이있으시면 dingpong@hitel.net 으로 메일부탁드립니다."
Label2.Caption = "                                              시나리오" + Chr(13) + Chr(13) _
                + "주인공이 쿠데타를 일으킨다....하지만 도와주는 사람없이 쓸쓸히 싸운다" + Chr(13) + Chr(13) _
                + "이젠 적이 돼버린 비행기를 부수고...결국 왕이 타고있는 비행기까지 부수게 돼면" + Chr(13) + Chr(13) _
                + "게임은 당신이 승리하게 되는 것이다 하지만..홀로 싸우기 때문에 위아래 좌우에서" + Chr(13) + Chr(13) _
                + "쏟아지는 적과 위에서 쏟아지는 미사일을 피해 왕까지 죽이는 일은 쉬운게 아니다." + Chr(13) + Chr(13) _
                + "만약 그것에 자신의 비행기 3대가 모두 격추당한다면 당신은 지게되는 것이다." + Chr(13) + Chr(13) _
                + "행운을 빈다.."
Label3.Caption = ": 이것을 먹으면 필살기를 다시 쓸수있게된다."
Label4.Caption = ": 게임화면 오른쪽아래 나타나는 것이다." + Chr(13) + Chr(13) _
                + "이것이 있으면 필살기를 쓸수있다."
End Sub

