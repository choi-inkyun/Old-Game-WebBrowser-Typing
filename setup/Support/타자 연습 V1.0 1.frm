VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�꼺��"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "Ÿ�� ���� V1.0 1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10005
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����"
      Height          =   495
      Left            =   7800
      Style           =   1  '�׷���
      TabIndex        =   26
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���� Ÿ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Style           =   1  '�׷���
      TabIndex        =   25
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ѱ� Ÿ��"
      BeginProperty Font 
         Name            =   "����"
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
      Style           =   1  '�׷���
      TabIndex        =   24
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���� ����"
      Height          =   495
      Left            =   2280
      Style           =   1  '�׷���
      TabIndex        =   19
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "���̵� ����"
      Height          =   1095
      Left            =   7680
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "�ʱ���"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�߱���"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�޴�(E)"
      Height          =   495
      Left            =   3960
      Style           =   1  '�׷���
      TabIndex        =   12
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�ߴ��ϱ�(&P)"
      Height          =   495
      Left            =   6000
      Style           =   1  '�׷���
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�����ϱ�(&S)"
      Height          =   495
      Left            =   6000
      MaskColor       =   &H00FFC0FF&
      Style           =   1  '�׷���
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
         Name            =   "����"
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
      Alignment       =   2  '��� ����
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
      Caption         =   "�Է� :"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�̸� :"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "���� ���� : "
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���� ���� : "
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
Dim ����(200), i, ���᰹��(10), ����(10), �̸�(10), ����(10), ���̵�(10), cnt, J, ����(100), ����(10)

Private Sub Command1_Click()
Randomize
Label4.Caption = 0
Label5.Caption = 10
For i = 0 To 6
If Command5.Enabled = False Then
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
ElseIf Command6.Enabled = False Then
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
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
MsgBox "Ÿ�� ������ ������ �ּ���." + Chr(13) + Chr(13) _
       + "�� �ؿ� �ѱ۰� ������ ������ �ּ���", vbCritical
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
Label1(i).Caption = ����(a)

a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
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
          PLAYERNAME = InputBox$("�̸��� �Է��ϼ���")
          If PLAYERNAME = "" Then PLAYERNAME = "�̸�����"
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
     MESSAGE = "���� ����" + Chr$(13)
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
'���᰹��(cnt) = Label4.Caption
'�̸�(cnt) = Text2.Text
'����(cnt) = ���᰹��(cnt) * 10
'If Command5.Enabled = True Then
'����(cnt) = "����"
'ElseIf Command6.Enabled = True Then
'����(cnt) = "�ѱ�"
'End If
'If Option1 = True Then
'���̵�(cnt) = "�����"
'ElseIf Option2 = True Then
'���̵�(cnt) = "�߱���"
'ElseIf Option3 = True Then
'���̵�(cnt) = "�ʱ���"
'Else
'���̵�(cnt) = "����"
'End If
'Form2.Show
'Form2.Print "����    �̸�    ���᰹��     ����    ���̵�     ����"

'For i = 1 To cnt
'����(i) = 1
'Next

'For i = 1 To cnt
'For J = 1 To cnt
'If ����(i) > ����(J) Then
'    ����(J) = ����(J) + 1
'End If
'Next
'Next


'For i = 1 To cnt - 1
'For J = i + 1 To cnt
'If ���᰹��(i) < ���᰹��(J) Then
'   im = ����(i)
'   ����(i) = ����(J)
'   ����(J) = im
   
 '  im = �̸�(i)
 '  �̸�(i) = �̸�(J)
 '  �̸�(J) = im
 '
 '  im = ���᰹��(i)
 '  ���᰹��(i) = ���᰹��(J)
 '  ���᰹��(J) = im
 '
 '  im = ����(i)
 '  ����(i) = ����(J)
 '  ����(J) = im
 '
 '  im = ���̵�(i)
 '  ���̵�(i) = ���̵�(J)
 '  ���̵�(J) = im
 '
 '  im = ����(i)
 '  ����(i) = ����(J)
 '  ����(J) = im
 '  End If
 '  Next
 '  Next

'For i = 1 To cnt
'Form2.Print Tab(2); ����(i);
'Form2.Print Tab(7); �̸�(i);
'Form2.Print Tab(16); ���᰹��(i);
'Form2.Print Tab(26); ����(i);
'Form2.Print Tab(32); ���̵�(i);
'Form2.Print Tab(42); ����(i)
'Next

End Sub

Private Sub Command5_Click()
Randomize
Command1.Enabled = True
Command2.Enabled = False
Text1.IMEMode = vbIMEModeHangul '�ѱ�
����(1) = "��ǻ��"
����(2) = "���ѹα�"
����(3) = "�̳���"
����(4) = "Ÿ��"
����(5) = "����"
����(6) = "��������"
����(7) = "�۰�"
����(8) = "�ػ�"
����(9) = "������"
����(10) = "��"
����(11) = "�Ͽ���"
����(12) = "Ű����"
����(13) = "������"
����(14) = "��ĳ��"
����(15) = "�޺�"
����(16) = "�౸"
����(17) = "��"
����(18) = "�豸"
����(19) = "��ġ"
����(20) = "Ź��"
����(21) = "�Ǳ�"
����(22) = "�ڵ庼"
����(23) = "���ǻ�"
����(24) = "�ݰ���"
����(25) = "�̱�"
����(26) = "��"
����(27) = "�丶��"
����(28) = "����"
����(29) = "����"
����(30) = "�ð�"
����(31) = "�Ͼ��"
����(32) = "����"
����(33) = "����"
����(34) = "������"
����(35) = "����"
����(36) = "����"
����(37) = "����"
����(38) = "���콺"
����(39) = "����Ŀ"
����(40) = "�׷���"
����(41) = "�ü�"
����(42) = "ĥ��"
����(43) = "���п���"
����(44) = "�ٶ�"
����(45) = "��"
����(46) = "���"
����(47) = "��"
����(48) = "��"
����(49) = "��"
����(50) = "��"
����(51) = "���"
����(52) = "�븣����"
����(53) = "�뱸"
����(54) = "��"
����(55) = "����"
����(56) = "���۽���"
����(57) = "���"
����(58) = "��"
����(59) = "õ��õ��"
����(60) = "Ȳ����"
����(61) = "�����"
����(62) = "�츲"
����(63) = "��ȭ"
����(64) = "��"
����(65) = "å��"
����(66) = "ĥ��"
����(67) = "����"
����(68) = "������"
����(69) = "��å"
����(70) = "��Ӵ�"
����(71) = "�ƹ���"
����(72) = "����"
����(73) = "���"
����(74) = "�̸�"
����(75) = "�ҸӴ�"
����(76) = "�Ҿƹ���"
����(77) = "����"
����(78) = "����"
����(79) = "����ö"
����(80) = "����"
����(81) = "����"
����(82) = "�ұ�"
����(83) = "��Ʈ��"
����(84) = "��"
����(85) = "�Ĺ�"
����(86) = "����"
����(87) = "�����"
����(88) = "����"
����(89) = "ȣ����"
����(90) = "���̿���"
����(91) = "���ͳ�"
����(92) = "��Ʈ��ũ"
����(93) = "����"
����(94) = "������"
����(95) = "��ġ"
����(96) = "��"
����(97) = "��ȭ"
����(98) = "�Ƿ���"
����(99) = "���̿ø�"
����(100) = "��"
����(101) = "�Ȱ�"
����(102) = "��"
����(103) = "������"
����(104) = "�Ҽ�"
����(105) = "å"
����(106) = "��ȭ"
����(107) = "����"
����(108) = "����"
����(109) = "�ɻ�"
����(110) = "����"
����(111) = "����"
����(112) = "����"
����(113) = "�ڰ���"
����(114) = "���̰�"
����(115) = "����"
����(116) = "��ȸ"
����(117) = "�ڷ�����"
����(118) = "�����´�"
����(119) = "�ƽþ�"
����(120) = "����"
����(121) = "�ֱ���"
����(122) = "�ϴ���"
����(123) = "����ȭ"
����(124) = "��"
����(125) = "��������"
����(126) = "Ȳ���ϴ�"
����(127) = "��"
����(128) = "����"
����(129) = "�״�"
����(130) = "����"
����(131) = "����"
����(132) = "����"
����(133) = "�ƺ�"
����(134) = "���"
����(135) = "����"
����(136) = "��帧"
����(137) = "��"
����(138) = "��"
����(139) = "���޶���"
����(140) = "���̾�"
����(141) = "ī��"
����(142) = "��"
����(143) = "����"
����(144) = "����"
����(145) = "�ܿ�"
����(146) = "�����ߵ�"
����(147) = "��"
����(148) = "��"
����(149) = "�߼�"
����(150) = "��ħ"
����(151) = "����"
����(152) = "����"
����(153) = "��"
����(154) = "����"
����(155) = "����"
����(156) = "�Ź�"
����(157) = "�ȭ"
����(158) = "������"
����(159) = "����"
����(160) = "���Ƹ�"
����(161) = "��ȣȸ"
����(162) = "����"
����(163) = "Į"
����(164) = "����"
����(165) = "���"
����(166) = "�ϴ�"
����(167) = "���"
����(168) = "����ϴ�"
����(169) = "��ȥ"
����(170) = "����"
����(171) = "����"
����(172) = "�Ұ���"
����(173) = "�з��Ͼ�"
����(174) = "����"
����(175) = "���"
����(176) = "�����Ǵ�"
����(177) = "��������"
����(178) = "����"
����(179) = "��"
����(180) = "�ձ�"
����(181) = "����"
����(182) = "���"
����(183) = "��"
����(184) = "�ܷӴ�"
����(185) = "�ܰ���"
����(186) = "����"
����(187) = "�칰"
����(188) = "����"
����(189) = "��θӸ�"
����(190) = "����"
����(191) = "�"
����(192) = "���θ�"
����(193) = "���ٰ�"
����(194) = "����"
����(195) = "����"
����(196) = "����"
����(197) = "����"
����(198) = "����"
����(199) = "����"
����(200) = "�Ѱᰰ��"


For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Label10.Caption = "1�ܰ�"
End Sub

Private Sub Command6_Click()
Randomize
Command1.Enabled = True
Command2.Enabled = False
Text1.IMEMode = vbIMEModeAlpha  '����
����(1) = "computer"
����(2) = "play"
����(3) = "can"
����(4) = "mouse"
����(5) = "key"
����(6) = "english"
����(7) = "knight"
����(8) = "korea"
����(9) = "man"
����(10) = "fool"
����(11) = "say"
����(12) = "star"
����(13) = "to"
����(14) = "enjoy"
����(15) = "car"
����(16) = "good"
����(17) = "baskstball"
����(18) = "team"
����(19) = "time"
����(20) = "join"
����(21) = "soccer"
����(22) = "save"
����(23) = "load"
����(24) = "prefer"
����(25) = "love"
����(26) = "like"
����(27) = "are"
����(28) = "do"
����(29) = "well"
����(30) = "kill"
����(31) = "king"
����(32) = "key"
����(33) = "short"
����(34) = "between"
����(35) = "start"
����(36) = "shout"
����(37) = "cross"
����(38) = "finger"
����(39) = "wish"
����(40) = "luck"
����(41) = "same"
����(42) = "tie"
����(43) = "saveral"
����(44) = "point"
����(45) = "stand"
����(46) = "both"
����(47) = "lose"
����(48) = "win"
����(49) = "because"
����(50) = "get"
����(51) = "one"
����(52) = "best"
����(53) = "country"
����(54) = "anyway"
����(55) = "there"
����(56) = "between"
����(57) = "go"
����(58) = "nothing"
����(59) = "would"
����(60) = "have"
����(61) = "cold"
����(62) = "that"
����(63) = "shall"
����(64) = "sounds"
����(65) = "see"
����(66) = "you"
����(67) = "then"
����(68) = "invite"
����(69) = "present"
����(70) = "send"
����(71) = "doctor"
����(72) = "glass"
����(73) = "bed"
����(74) = "rest"
����(75) = "roof"
����(76) = "porch"
����(77) = "walk"
����(78) = "into"
����(79) = "jump"
����(80) = "full"
����(81) = "wonderful"
����(82) = "unhappy"
����(83) = "happy"
����(84) = "week"
����(85) = "much"
����(86) = "what"
����(87) = "movi"
����(88) = "rise"
����(89) = "body"
����(90) = "age"
����(91) = "egg"
����(92) = "hour"
����(93) = "also"
����(94) = "clear"
����(95) = "healthy"
����(96) = "look"
����(97) = "matter"
����(98) = "mother"
����(99) = "father"
����(100) = "easy"

Label10.Caption = "1�ܰ�"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
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
MsgBox "< �̽��� ���� >" + Chr(13) + Chr(13) _
       + "�̽��� ���� ������ ���α׷� �Դϴ�." + Chr(13) + Chr(13) _
       + "ó���� �����Ҷ� �� �ؿ��ִ� �ѱ۰� �����߿��� ������ �ּ���." + Chr(13) + Chr(13) _
       + "���̵� ������ ���ּ���. �׸��� ���α׷��� ���� ������ �������ϴ�." + Chr(13) + Chr(13) _
       + "��������� ������ ������ ������. ������ ��� �������� �����Դϴ�." + Chr(13) + Chr(13) _
       + "���᰹���� ���� 11�ܰ�� ������ ���ϴ�." + Chr(13) + Chr(13) _
       + "��̰� �ϼ���..^^" + Chr(13) + Chr(13)

End Sub

Private Sub Form_Load()
Command1.Enabled = True
Command2.Enabled = False
Command5.Enabled = True
MsgBox "������ ������ �꼺�� ����ִ� ���� �������� �ִ�" + Chr(13) + Chr(13) _
        + "�츮�� �꼺�� ���ƾ� �Ѵ�."
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
           Label1(i).Caption = ����(a)
           If Label1(i).Caption = "" Then
           a = Int(Rnd(1) * 200)
           Label1(i).Caption = ����(a)
           End If
           ElseIf Command6.Enabled = False Then
             a = Int(Rnd(1) * 100)
           Label1(i).Caption = ����(a)
          
End If
           
            End If
              Next
    Text1.Text = ""
    Text1.SetFocus

If Command5.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2�ܰ�"
Label10.Caption = "2�ܰ�"
For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3�ܰ�"
Label10.Caption = "3�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4�ܰ�"
Label10.Caption = "4�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5�ܰ�"
Label10.Caption = "5�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6�ܰ�"
Label10.Caption = "6�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7�ܰ�"
Label10.Caption = "7�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8�ܰ�"
Label10.Caption = "8�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9�ܰ�"
Label10.Caption = "9�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10�ܰ�"
Label10.Caption = "10�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "������ �ܰ�"
Label10.Caption = "������ �ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
End If
End If



If Command6.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2�ܰ�"
Label10.Caption = "2�ܰ�"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3�ܰ�"
Label10.Caption = "3�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4�ܰ�"
Label10.Caption = "4�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5�ܰ�"
Label10.Caption = "5�ܰ�"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6�ܰ�"
Label10.Caption = "6�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7�ܰ�"
Label10.Caption = "7�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8�ܰ�"
Label10.Caption = "8�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9�ܰ�"
Label10.Caption = "9�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10�ܰ�"
Label10.Caption = "10�ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "������ �ܰ�"
Label10.Caption = "������ �ܰ�"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = ����(a)
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
If Label10.Caption = "1�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 25)
ElseIf Label10.Caption = "2�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 30)
ElseIf Label10.Caption = "3�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "4�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 35)
ElseIf Label10.Caption = "5�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 40)
ElseIf Label10.Caption = "6�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 40)
ElseIf Label10.Caption = "7�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 45)
ElseIf Label10.Caption = "8�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 45)
ElseIf Label10.Caption = "9�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 45)
ElseIf Label10.Caption = "10�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 50)
ElseIf Label10.Caption = "������ �ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 55)
End If
ElseIf Option2 = True Then
If Label10.Caption = "1�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 15)
ElseIf Label10.Caption = "2�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "3�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 20)
ElseIf Label10.Caption = "4�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 25)
ElseIf Label10.Caption = "5�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "6�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 30)
ElseIf Label10.Caption = "7�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(22) * 32)
ElseIf Label10.Caption = "8�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 35)
ElseIf Label10.Caption = "9�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(28) * 38)
ElseIf Label10.Caption = "10�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 40)
ElseIf Label10.Caption = "������ �ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 45)
End If
ElseIf Option3 = True Then
If Label10.Caption = "1�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(1) * 5)
ElseIf Label10.Caption = "2�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 10)
ElseIf Label10.Caption = "3�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 15)
ElseIf Label10.Caption = "4�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "5�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 25)
ElseIf Label10.Caption = "6�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 30)
ElseIf Label10.Caption = "7�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 32)
ElseIf Label10.Caption = "8�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 34)
ElseIf Label10.Caption = "9�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 38)
ElseIf Label10.Caption = "10�ܰ�" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 38)
ElseIf Label10.Caption = "������ �ܰ�" Then
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
            Label1(i).Caption = ����(a)
            ElseIf Command6.Enabled = False Then
            a = Int(Rnd(1) * 100)
            Label1(i).Caption = ����(a)
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
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = ����(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
End If

If Command6.Enabled = False Then
If Label1(0).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = ����(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = ����(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = ����(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = ����(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = ����(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = ����(a)
End If
End If


End Sub

Private Sub Timer2_Timer()


End Sub
