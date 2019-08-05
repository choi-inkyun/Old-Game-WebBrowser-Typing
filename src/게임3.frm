VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  '없음
   Caption         =   "Form4"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   Icon            =   "게임3.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "게임3.frx":030A
   ScaleHeight     =   7020
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   0
      Picture         =   "게임3.frx":0614
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   -120
      Width           =   9615
      Begin MCI.MMControl midi2 
         Height          =   330
         Left            =   3360
         TabIndex        =   4
         Top             =   2640
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "도움말"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   5280
         TabIndex        =   5
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "순위 보기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "종료 하기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "게임 시작"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   5400
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
midi2.Visible = False
midi2.FileName = App.Path + "\altkdlf.wav"
midi2.Command = "open"
midi2.Command = "play"
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub

Private Sub Label1_Click()
Form6.Show

Unload Me

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF00&

End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF00&
End Sub

Private Sub Label3_Click()
     ReDim PLAYER(1 To 11) As String
     ReDim SCORE(1 To 11) As Integer
     Dim PLAYERNAME As String
     Dim PLAYERTEMP As String
     Dim SCORETEMP As Integer
     Dim MAINLOOP As Integer
     Dim LOOPCTR As Integer
     Dim FOUNDSW As Integer
     Open App.Path & "\game.txt" For Input As #1
     For LOOPCTR = 1 To 10
          Input #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     For LOOPCTR = 1 To 10
          If wjatn > SCORE(LOOPCTR) Then FOUNDSW = 1
     Next LOOPCTR
     
     
     
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
     Open "game.txt" For Output As #1
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

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label3.ForeColor = &HFF00&

End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)

End Sub

Private Sub Label4_Click()
Form5.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF00&

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFF&
Label2.ForeColor = &HFF&
Label3.ForeColor = &HFF&
Label4.ForeColor = &HFFFF&

End Sub
