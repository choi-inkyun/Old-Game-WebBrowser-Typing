VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form7 
   Caption         =   "����"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   9480
   Icon            =   "Ÿ�� ���� V1.0 7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   6855
   ScaleWidth      =   9480
   StartUpPosition =   2  'ȭ�� ���
   Begin MCI.MMControl mid12 
      Height          =   330
      Left            =   1680
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl mid11 
      Height          =   330
      Left            =   1800
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer16 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  '�׷���
      TabIndex        =   32
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer Timer14 
      Interval        =   100
      Left            =   5400
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�ߴ��ϱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  '�׷���
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer Timer13 
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   4680
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "���̵�"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   1695
      Begin VB.OptionButton Option3 
         Caption         =   "�ʺ���"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�߱���"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�޴���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6120
      Style           =   1  '�׷���
      TabIndex        =   10
      Top             =   6120
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ó������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4200
      Style           =   1  '�׷���
      TabIndex        =   9
      Top             =   6120
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�����ϱ�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   360
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   6120
      Width           =   1410
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.Label Label13 
      Height          =   420
      Index           =   1
      Left            =   -600
      TabIndex        =   34
      Top             =   2640
      Width           =   1020
   End
   Begin VB.Label Label13 
      Height          =   420
      Index           =   0
      Left            =   -600
      TabIndex        =   33
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label11 
      Height          =   975
      Left            =   8040
      TabIndex        =   30
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "������ ����"
      Height          =   255
      Left            =   8160
      TabIndex        =   29
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   11
      Left            =   7800
      TabIndex        =   27
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   10
      Left            =   7800
      TabIndex        =   26
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   9
      Left            =   7800
      TabIndex        =   25
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   24
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Ȯ �� :"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Ȯ �� :"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   7200
      TabIndex        =   16
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   15
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "���� :"
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
      Left            =   6240
      TabIndex        =   14
      Top             =   5280
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      X1              =   0
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      Caption         =   "�� �� :"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   3960
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '�ܻ�
      Height          =   3975
      Left            =   720
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   7
      Left            =   7800
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   6
      Left            =   7800
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ����(200), k, g, e, p, o, i, u, t, r, m, n, s, q, w, v, f, h, J, l, aa, bb, cc, dd, aaa, bbb, ccc, ddd
Dim b As Integer
Dim qwe
Dim ttt
Dim tttt
Dim yyy
Dim qqq
Dim www
Dim eee
Dim rrr
Dim ytr
Private Sub Command1_Click()
Text1.IMEMode = vbIMEModeHangul '�ѱ�
Timer1.Enabled = True
Label2.Visible = True
Text1.SetFocus
Command1.Enabled = False
Command2.Enabled = True
Frame1.Visible = False
Frame1.Enabled = False
p = False
q = False
w = False
v = False
f = False
h = False
J = False
l = False
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
Timer10.Enabled = True
Timer11.Enabled = True
Timer12.Enabled = True
Timer13.Enabled = True
Timer14.Enabled = True
Timer15.Enabled = True
Timer16.Enabled = True
Command4.Enabled = True
Label8.Visible = True
Label6.Visible = True
Label12.Visible = True
Label11.Caption = "���� ���� ������ ���۵Ǿ����ϴ�. ���� �����Ű�°� �����Դϴ�. ����� ���ڽ��ϴ�."
End Sub

Private Sub Command2_Click()
Label11.Caption = "������ �ٽ� �߹��Ϸ��� �մϴ�"
Command1.Enabled = True
Command2.Enabled = False
Shape1.Left = 800
Timer1.Enabled = False
Label8.Visible = False
Label6.Visible = False
Label5.Visible = 0
Frame1.Visible = True
Frame1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Command4.Enabled = False
Label12.Visible = False
Label2.Visible = False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
Frame1.Visible = True
Frame1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label11.Caption = "�����Դϴ�"
Label2.Visible = False
End Sub

Private Sub Command5_Click()
MsgBox "<���� ����>" + Chr(13) + Chr(13) _
      + "�ȳ��ϼ���. ���￡���� ������ ������ 0 ���� ����� �̱�� �˴ϴ�" + Chr(13) + Chr(13) _
      + "������ ������ ������ 7800������ �Ǹ� ���� �˴ϴ�" + Chr(13) + Chr(13) _
      + "�������� �ܾ ���ִ� �ܾ ġ�ø� �ܾ ������ �ο�� �˴ϴ�" + Chr(13) + Chr(13) _
      + "�� ���￡���� ���� ����� �¸��Ѱ��� �߿��մϴ�" + Chr(13) + Chr(13) _
      + "����� ������...����ְ� �ϼ���~.."
End Sub

Private Sub Form_Load()
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


a = Int(Rnd(1) * 200)
Label2.Caption = ����(a)
Label5.Caption = 0
Label2.Visible = False
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label12.Visible = False
mid11.FileName = App.Path + "\THEME MI.wav"
mid11.Command = "open"
mid12.FileName = App.Path + "\THEME AS.wav"
mid12.Command = "open"
For i = 0 To 1
Label13(i).Visible = False
Next
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
Label8.Caption = Text1.Text
'Label8.Caption = tirm(Text1.Text)
'p = Text1.Text
        ttt = Len(Text1.Text)
          tttt = ttt * 40
If Trim(Text1.Text) = Label2.Caption Then
    c = Int(Rnd(0) * 11)
    If c = hh Then
        c = Int(Rnd(0) * 11)
            If c = hh Then
            ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
            ElseIf c = hh Then
        c = Int(Rnd(0) * 11)

        c = Int(Rnd(0) * 11)
    If c = hh Then
        c = Int(Rnd(0) * 11)

    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
   End If
   End If
   End If
    hh = c
    s = 0
    o = 1
    i = 2
    u = 3
    t = 4
    r = 5
    n = 6
    m = 7
    aa = 8
    bb = 9
    cc = 10
    dd = 11
    Label1(c).Caption = Label2.Caption
   
    a = Int(Rnd(1) * 200)
    Label2.Caption = ����(a)
    If Label2.Caption = "" Then
    a = Int(Rnd(1) * 200)
    Label2.Caption = ����(a)
     End If
    g = k * 80
    Label6.Caption = "�¾Ҵ�"
If c = s Then
p = True
End If
If c = m Then
q = True
End If
If c = n Then
w = True
End If
If c = r Then
v = True
End If
If c = t Then
f = True
End If
If c = u Then
h = True
End If
If c = i Then
J = True
End If
If c = o Then
l = True
End If
If c = aa Then
aaa = True
End If
If c = bb Then
bbb = True
End If
If c = cc Then
ccc = True
End If
If d = dd Then
ddd = True
End If
        Else
If Text1.Text <> "" Then
    yyy = Int(Rnd(0) * 1)
    qqq = 0
    eee = 1
    ytr = yyy
    If yyy = ytr Then
     yyy = Int(Rnd(0) * 1)

    ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
  End If
    Label13(yyy).Caption = Text1.Text
    If yyy = qqq Then
    www = True
    End If
    If yyy = eee Then
    rrr = True
    End If
    a = Int(Rnd(1) * 200)
    Label2.Caption = ����(a)
    If Label2.Caption = "" Then
     a = Int(Rnd(1) * 200)
     Label2.Caption = ����(a)
    End If
    Label6.Caption = "Ʋ�ȴ�"
End If
End If
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()

If Option1.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(65) * 30)
ElseIf Option2.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(50) * 20)
ElseIf Option3.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(25) * 5)
Else
Shape1.Left = Shape1.Left + Int(Rnd(30) * 20)
End If
e = Label5.Caption
k = Len(Label2.Caption)

If Shape1.Left >= 7800 Then
MsgBox "������ �������ϴ�"
MsgBox "����� �� �ο� �־����� ���￡���� ���ϰ� ���ҽ��ϴ�." + Chr(13) + Chr(13) _
       + "����� �������� " & e & " ��ŭ�� ���ظ� �������ϴ�"
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label1(0).Caption = ""
Label1(0).Left = 7800
Label1(1).Caption = ""
Label1(1).Left = 7800
Label1(2).Caption = ""
Label1(2).Left = 7800
Label1(3).Caption = ""
Label1(3).Left = 7800
Label1(4).Caption = ""
Label1(4).Left = 7800
Label1(5).Caption = ""
Label1(5).Left = 7800
Label1(6).Caption = ""
Label1(6).Left = 7800
Label1(7).Caption = ""
Label1(7).Left = 7800
Label1(8).Caption = ""
Label1(8).Left = 7800
Label1(9).Caption = ""
Label1(9).Left = 7800
Label1(10).Caption = ""
Label1(10).Left = 7800
Label1(11).Caption = ""
Label1(11).Left = 7800
Label13(0).Caption = ""
Label13(0).Left = 0
Label13(1).Caption = ""
Label13(1).Left = 0
Label12.Caption = ""
Text1.Text = ""
Command1.SetFocus
Shape1.Left = 840
Label5.Caption = 0
Label2.Visible = False
Frame1.Visible = True
Frame1.Enabled = True
Label11.Caption = "����� ���￡�� �й��Ͽ����ϴ�"
End If

If Shape1.Left <= 0 Then
MsgBox "�����մϴ�.^^"
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label1(0).Caption = ""
Label1(0).Left = 7800
Label1(1).Caption = ""
Label1(1).Left = 7800
Label1(2).Caption = ""
Label1(2).Left = 7800
Label1(3).Caption = ""
Label1(3).Left = 7800
Label1(4).Caption = ""
Label1(4).Left = 7800
Label1(5).Caption = ""
Label1(5).Left = 7800
Label1(6).Caption = ""
Label1(6).Left = 7800
Label1(7).Caption = ""
Label1(7).Left = 7800
Label1(8).Caption = ""
Label1(8).Left = 7800
Label1(9).Caption = ""
Label1(9).Left = 7800
Label1(10).Caption = ""
Label1(10).Left = 7800
Label1(11).Caption = ""
Label1(11).Left = 7800
Label13(0).Caption = ""
Label13(0).Left = 0
Label13(1).Caption = ""
Label13(1).Left = 0
Label12.Caption = ""
Text1.Text = ""
Command1.SetFocus
Shape1.Left = 840
MsgBox "������ �������ϴ�"
MsgBox "�����մϴ�. ����� ���￡�� �¸��Ͽ����ϴ�" + Chr(13) + Chr(13) _
        + "����� " & e & "���� ����� �¸��ϼ̽��ϴ�"
Label5.Caption = 0
Label2.Visible = False
Frame1.Visible = True
Frame1.Enabled = True
Label11.Caption = "�����մϴ�. ����� ���￡�� �¸��Ͽ����ϴ�"
End If
End Sub

Private Sub Timer10_Timer()
Label1(aa).Visible = False

If aaa Then
           If Label1(aa).Left <= 7200 Then
       Label1(aa).Visible = True
        End If
    
    Label1(aa).Left = Label1(aa).Left - 800
    If Label1(aa).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
       mid11.Command = "prev"

       mid11.Command = "play"
               Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(aa).Left = 7800
        Label1(aa).Caption = ""
        aaa = False
    End If
End If

End Sub

Private Sub Timer11_Timer()
Label1(bb).Visible = False

If bbb Then
           If Label1(bb).Left <= 7200 Then
       Label1(bb).Visible = True
        End If
    
    Label1(bb).Left = Label1(bb).Left - 800
    If Label1(bb).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
       mid11.Command = "prev"
        mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(bb).Left = 7800
        Label1(bb).Caption = ""
        bbb = False
    End If
End If

End Sub

Private Sub Timer12_Timer()
Label1(cc).Visible = False

If ccc Then
           If Label1(cc).Left <= 7200 Then
       Label1(cc).Visible = True
        End If
    
    Label1(cc).Left = Label1(cc).Left - 800
    If Label1(cc).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(cc).Left = 7800
        Label1(cc).Caption = ""
        ccc = False
    End If
End If

End Sub

Private Sub Timer13_Timer()
Label1(dd).Visible = False

If ddd Then
           If Label1(dd).Left <= 7200 Then
       Label1(dd).Visible = True
        End If
    
    Label1(dd).Left = Label1(dd).Left - 800
    If Label1(dd).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(dd).Left = 7800
        Label1(dd).Caption = ""
        ddd = False
    End If
End If

End Sub

Private Sub Timer14_Timer()
Label12.Caption = Shape1.Left
End Sub

Private Sub Timer15_Timer()
Label13(qqq).Visible = False
If www Then
    Label13(qqq).Visible = True
    Label13(qqq).Left = Label13(qqq).Left + 1000
    If Label13(qqq).Left >= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left + tttt
mid12.Command = "prev"
        
        mid12.Command = "play"

       Label11.Caption = "�������� " & tttt & "���� ������ �����߽��ϴ�"
         Label13(qqq).Caption = Text1.Text
        Shape1.Left = Shape1.Left + tttt
       Label5.Caption = Label5.Caption - tttt
        Label13(qqq).Left = 0
        Label13(qqq).Caption = ""
        www = False
    End If
End If
End Sub

Private Sub Timer16_Timer()
Label13(eee).Visible = False
If rrr Then
    Label13(eee).Visible = True
    Label13(eee).Left = Label13(eee).Left + 1000
    If Label13(eee).Left >= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left + tttt
mid12.Command = "prev"
        
        mid12.Command = "play"

       Label11.Caption = "�������� " & tttt & "���� ������ �����߽��ϴ�"
         Label13(eee).Caption = Text1.Text
        Shape1.Left = Shape1.Left + tttt
       Label5.Caption = Label5.Caption - tttt
        Label13(eee).Left = 0
        Label13(eee).Caption = ""
        rrr = False
    End If
End If

End Sub

Private Sub Timer2_Timer()
Label1(o).Visible = False
If l Then
    If Label1(o).Left <= 7200 Then
       Label1(o).Visible = True
       End If
    
    Label1(o).Left = Label1(o).Left - 800
    If Label1(o).Left <= (Shape1.Left + Shape1.Width) Then
      Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

       Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(o).Left = 7800
        Label1(o).Caption = ""
        l = False
    End If
End If

End Sub

Private Sub Timer3_Timer()
Label1(i).Visible = False

If J Then
           If Label1(i).Left <= 7200 Then
       Label1(i).Visible = True
      End If
    
    Label1(i).Left = Label1(i).Left - 800
    If Label1(i).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(i).Left = 7800
        Label1(i).Caption = ""
        J = False
    End If
End If

End Sub

Private Sub Timer4_Timer()
Label1(u).Visible = False

If h Then
           If Label1(u).Left <= 7200 Then
       Label1(u).Visible = True
        End If
    
    Label1(u).Left = Label1(u).Left - 800
    If Label1(u).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(u).Left = 7800
        Label1(u).Caption = ""
        h = False
    End If
End If

End Sub

Private Sub Timer5_Timer()
Label1(t).Visible = False

If f Then
           If Label1(t).Left <= 7200 Then
       Label1(t).Visible = True
       End If
    
    Label1(t).Left = Label1(t).Left - 800
    If Label1(t).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(t).Left = 7800
        Label1(t).Caption = ""
        f = False
    End If
End If

End Sub

Private Sub Timer6_Timer()
Label1(r).Visible = False
If v Then
           If Label1(r).Left <= 7200 Then
       Label1(r).Visible = True
         End If
     
    Label1(r).Left = Label1(r).Left - 800
    If Label1(r).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(r).Left = 7800
        Label1(r).Caption = ""
        v = False
    End If
End If

End Sub

Private Sub Timer7_Timer()
Label1(n).Visible = False

If w Then
           If Label1(n).Left <= 7200 Then
       Label1(n).Visible = True
       End If
    Label1(n).Left = Label1(n).Left - 800
        If Label1(n).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(n).Left = 7800
        Label1(n).Caption = ""
        w = False
    End If
End If

End Sub

Private Sub Timer8_Timer()
Label1(m).Visible = False

If q Then
           If Label1(m).Left <= 7200 Then
       Label1(m).Visible = True
        End If
    
    Label1(m).Left = Label1(m).Left - 800
    If Label1(m).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
        Label1(m).Left = 7800
        Label1(m).Caption = ""
            q = False
    End If
End If

End Sub

Private Sub Timer9_Timer()
Label1(s).Visible = False

If p Then
                 If Label1(s).Left <= 7200 Then
       Label1(s).Visible = True
        End If
    
    Label1(s).Left = Label1(s).Left - 800
    If Label1(s).Left <= (Shape1.Left + Shape1.Width) Then

       Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "������ ���縦 " & g & " �� ��ŭ �սǽ��׽��ϴ�"
       Label1(s).Left = 7800
        Label1(s).Caption = ""
        p = False
    End If
End If

End Sub
