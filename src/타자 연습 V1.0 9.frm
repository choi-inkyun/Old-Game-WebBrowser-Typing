VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "�޸���"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "Ÿ�� ���� V1.0 9.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   7365
   ScaleWidth      =   8805
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame2 
      Caption         =   "�����"
      Height          =   1215
      Left            =   2400
      TabIndex        =   27
      Top             =   6000
      Width           =   4215
      Begin VB.Label Label26 
         Alignment       =   2  '��� ����
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   2  '��� ����
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   2  '��� ����
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  '��� ����
         Caption         =   "99"
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  '��� ����
         Caption         =   "99"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  '��� ����
         Caption         =   "99"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  '��� ����
         Caption         =   "�ʱ���"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   2  '��� ����
         Caption         =   "�߱���"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  '��� ����
         Caption         =   "�����"
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
      Caption         =   "���̵�"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "�ʱ���"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�߱���"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����"
      Height          =   495
      Left            =   7320
      Style           =   1  '�׷���
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
      Caption         =   "ó������"
      Height          =   495
      Left            =   5640
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�޴�"
      Height          =   495
      Left            =   7320
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�����ϱ�(&S)"
      Height          =   495
      Left            =   5640
      Style           =   1  '�׷���
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
         Name            =   "����"
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
      Caption         =   "�̸� :"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Picture         =   "Ÿ�� ���� V1.0 9.frx":030A
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   120
      Picture         =   "Ÿ�� ���� V1.0 9.frx":0BD4
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "Ÿ�� ���� V1.0 9.frx":149E
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "Ÿ�� ���� V1.0 9.frx":1D68
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Ÿ�� ���� V1.0 9.frx":2632
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Ÿ�� ���� V1.0 9.frx":2EFC
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "�ܾ� :"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label1 
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
Dim ����(200), k, b, c, d, e, f, g, h, i
Dim cnt As Byte
Dim aaa1 As Byte
Dim aaa2 As Byte
Dim aaa3 As Byte
Dim nam
Private Sub Command1_Click()
Label9.Caption = "���ϸ� �޸��� ���Ⱑ �ְڽ��ϴ�."
Text1.IMEMode = vbIMEModeHangul '�ѱ�
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
Label1.Caption = ����(a)
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
Label9.Caption = "��Ⱑ ���۵Ǿ����ϴ�"
Timer8.Enabled = True
Frame1.Visible = False
Label10.Visible = True
Label10.Caption = Text2.Text
nam = Text2.Text
If Text2.Text = "" Then
nam = "��ī��"
Label10 = "��ī��"
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
MsgBox "<�޸��� ����>" + Chr(13) + Chr(13) _
       + "����ڴ� " & nam & "�� �����ϰ� �˴ϴ�." + Chr(13) + Chr(13) _
       + "�ܾ ���ִ� �ܾ ���߽ø� �˴ϴ�." + Chr(13) + Chr(13) _
       + "����ְ� �ϼ���..~.^^.."
End Sub

Private Sub Form_Load()
Label9.Caption = "���ϸ� ���� �޸��� ��ȸ�� �������ϴ�. ������ �غ�."
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
k = Len(����(a))
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
Label1.Caption = ����(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = ����(a)
End If
k = Len(����(a))
b = k * Int(Rnd(140) * 170)
Image6.Left = Image6.Left + b
Else
If Text1.Text <> "" Then
Image6.Left = Image6.Left - Int(Rnd(40) * 70)
a = Int(Rnd(1) * 200)
Label1.Caption = ����(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = ����(a)
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
Label3.Caption = "1��"
Label9.Caption = "�̻��ؾ��� 1���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = "�̻��ؾ�"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = "�̻��ؾ�"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = "�̻��ؾ�"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If


End If
End If
End If
End If
End If
If Label4.Caption = "1��" Then
Label3.Caption = "2��"
Label9.Caption = "�̻��ؾ��� 2���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label5.Caption = "1��" Then
Label3.Caption = "2��"
Label9.Caption = "�̻��ؾ��� 2���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label6.Caption = "1��" Then
Label3.Caption = "2��"
Label9.Caption = "�̻��ؾ��� 2���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label7.Caption = "1��" Then
Label3.Caption = "2��"
Label9.Caption = "�̻��ؾ��� 2���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label8.Caption = "1��" Then
Label3.Caption = "2��"
Label9.Caption = "�̻��ؾ��� 2���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label4.Caption = "2��" Then
Label3.Caption = "3��"
Label9.Caption = "�̻��ؾ��� 3���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label5.Caption = "2��" Then
Label3.Caption = "3��"
Label9.Caption = "�̻��ؾ��� 3���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label6.Caption = "2��" Then
Label3.Caption = "3��"
Label9.Caption = "�̻��ؾ��� 3���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label7.Caption = "2��" Then
Label3.Caption = "3��"
Label9.Caption = "�̻��ؾ��� 3���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label8.Caption = "2��" Then
Label3.Caption = "3��"
Label9.Caption = "�̻��ؾ��� 3���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label4.Caption = "3��" Then
Label3.Caption = "4��"
Label9.Caption = "�̻��ؾ��� 4���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label5.Caption = "3��" Then
Label3.Caption = "4��"
Label9.Caption = "�̻��ؾ��� 4���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label6.Caption = "3��" Then
Label3.Caption = "4��"
Label9.Caption = "�̻��ؾ��� 4���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label7.Caption = "3��" Then
Label3.Caption = "4��"
Label9.Caption = "�̻��ؾ��� 4���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label8.Caption = "3��" Then
Label3.Caption = "4��"
Label9.Caption = "�̻��ؾ��� 4���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label4.Caption = "4��" Then
Label3.Caption = "5��"
Label9.Caption = "�̻��ؾ��� 5���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label5.Caption = "4��" Then
Label3.Caption = "5��"
Label9.Caption = "�̻��ؾ��� 5���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label6.Caption = "4��" Then
Label3.Caption = "5��"
Label9.Caption = "�̻��ؾ��� 5���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label7.Caption = "4��" Then
Label3.Caption = "5��"
Label9.Caption = "�̻��ؾ��� 5���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label8.Caption = "4��" Then
Label3.Caption = "5��"
Label9.Caption = "�̻��ؾ��� 5���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label4.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "�̻��ؾ��� 6���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label5.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "�̻��ؾ��� 6���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label6.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "�̻��ؾ��� 6���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label7.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "�̻��ؾ��� 6���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
End If
If Label8.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "�̻��ؾ��� 6���� �Ͽ����ϴ�"
Label11.Caption = cnt & "��"
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
Label4.Caption = "1��"
Label9.Caption = "���̸��� 1���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = "���̸�"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = "���̸�"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = "���̸�"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1��" Then
Label4.Caption = "2��"
Label9.Caption = "���̸��� 2���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label5.Caption = "1��" Then
Label4.Caption = "2��"
Label9.Caption = "���̸��� 2���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label6.Caption = "1��" Then
Label4.Caption = "2��"
Label9.Caption = "���̸��� 2���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label7.Caption = "1��" Then
Label4.Caption = "2��"
Label9.Caption = "���̸��� 2���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label8.Caption = "1��" Then
Label4.Caption = "2��"
Label9.Caption = "���̸��� 2���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label3.Caption = "2��" Then
Label4.Caption = "3��"
Label9.Caption = "���̸��� 3���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label5.Caption = "2��" Then
Label4.Caption = "3��"
Label9.Caption = "���̸��� 3���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label6.Caption = "2��" Then
Label4.Caption = "3��"
Label9.Caption = "���̸��� 3���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label7.Caption = "2��" Then
Label4.Caption = "3��"
Label9.Caption = "���̸��� 3���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label8.Caption = "2��" Then
Label4.Caption = "3��"
Label9.Caption = "���̸��� 3���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label3.Caption = "3��" Then
Label4.Caption = "4��"
Label9.Caption = "���̸��� 4���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label5.Caption = "3��" Then
Label4.Caption = "4��"
Label9.Caption = "���̸��� 4���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label6.Caption = "3��" Then
Label4.Caption = "4��"
Label9.Caption = "���̸��� 4���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label7.Caption = "3��" Then
Label4.Caption = "4��"
Label9.Caption = "���̸��� 4���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label8.Caption = "3��" Then
Label4.Caption = "4��"
Label9.Caption = "���̸��� 4���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label3.Caption = "4��" Then
Label4.Caption = "5��"
Label9.Caption = "���̸��� 5���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label5.Caption = "4��" Then
Label4.Caption = "5��"
Label9.Caption = "���̸��� 5���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label6.Caption = "4��" Then
Label4.Caption = "5��"
Label9.Caption = "���̸��� 5���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label7.Caption = "4��" Then
Label4.Caption = "5��"
Label9.Caption = "���̸��� 5���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label8.Caption = "4��" Then
Label4.Caption = "5��"
Label9.Caption = "���̸��� 5���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label3.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "���̸��� 6���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label5.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "���̸��� 6���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label6.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "���̸��� 6���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label7.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "���̸��� 6���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
End If
If Label8.Caption = "5��" Then
Label4.Caption = "6��"
Label9.Caption = "���̸��� 6���� �Ͽ����ϴ�"
Label12.Caption = cnt & "��"
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
Label5.Caption = "1��"
Label9.Caption = "���αⰡ 1���� �Ͽ����ϴ�"
Label13.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = "���α�"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = "���α�"
End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = "���α�"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1��" Then
Label5.Caption = "2��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 2���� �Ͽ����ϴ�"
End If
If Label4.Caption = "1��" Then
Label5.Caption = "2��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 2���� �Ͽ����ϴ�"
End If
If Label6.Caption = "1��" Then
Label5.Caption = "2��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 2���� �Ͽ����ϴ�"
End If
If Label7.Caption = "1��" Then
Label5.Caption = "2��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 2���� �Ͽ����ϴ�"
End If
If Label8.Caption = "1��" Then
Label5.Caption = "2��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 2���� �Ͽ����ϴ�"
End If
If Label3.Caption = "2��" Then
Label5.Caption = "3��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 3���� �Ͽ����ϴ�"
End If
If Label4.Caption = "2��" Then
Label5.Caption = "3��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 3���� �Ͽ����ϴ�"
End If
If Label6.Caption = "2��" Then
Label5.Caption = "3��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 3���� �Ͽ����ϴ�"
End If
If Label7.Caption = "2��" Then
Label5.Caption = "3��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 3���� �Ͽ����ϴ�"
End If
If Label8.Caption = "2��" Then
Label5.Caption = "3��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 3���� �Ͽ����ϴ�"
End If
If Label3.Caption = "3��" Then
Label5.Caption = "4��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 4���� �Ͽ����ϴ�"
End If
If Label4.Caption = "3��" Then
Label5.Caption = "4��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 4���� �Ͽ����ϴ�"
End If
If Label6.Caption = "3��" Then
Label5.Caption = "4��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 4���� �Ͽ����ϴ�"
End If
If Label7.Caption = "3��" Then
Label5.Caption = "4��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 4���� �Ͽ����ϴ�"
End If
If Label8.Caption = "3��" Then
Label5.Caption = "4��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 4���� �Ͽ����ϴ�"
End If
If Label3.Caption = "4��" Then
Label5.Caption = "5��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 5���� �Ͽ����ϴ�"
End If
If Label4.Caption = "4��" Then
Label5.Caption = "5��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 5���� �Ͽ����ϴ�"
End If
If Label6.Caption = "4��" Then
Label5.Caption = "5��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 5���� �Ͽ����ϴ�"
End If
If Label7.Caption = "4��" Then
Label5.Caption = "5��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 5���� �Ͽ����ϴ�"
End If
If Label8.Caption = "4��" Then
Label5.Caption = "5��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 5���� �Ͽ����ϴ�"
End If
If Label3.Caption = "5��" Then
Label5.Caption = "6��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 6���� �Ͽ����ϴ�"
End If
If Label4.Caption = "5��" Then
Label5.Caption = "6��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 6���� �Ͽ����ϴ�"
End If
If Label6.Caption = "5��" Then
Label5.Caption = "6��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 6���� �Ͽ����ϴ�"
End If
If Label7.Caption = "5��" Then
Label5.Caption = "6��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 6���� �Ͽ����ϴ�"
End If
If Label8.Caption = "5��" Then
Label5.Caption = "6��"
Label13.Caption = cnt & "��"
Label9.Caption = "���α�� 6���� �Ͽ����ϴ�"
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
Label6.Caption = "1��"
Label9.Caption = "�Ḹ���� 1���� �Ͽ����ϴ�"
Label14.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = "�Ḹ��"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = "�Ḹ��"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = "�Ḹ��"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1��" Then
Label6.Caption = "2��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 2���� �Ͽ����ϴ�"
End If
If Label4.Caption = "1��" Then
Label6.Caption = "2��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 2���� �Ͽ����ϴ�"
End If
If Label5.Caption = "1��" Then
Label6.Caption = "2��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 2���� �Ͽ����ϴ�"
End If
If Label7.Caption = "1��" Then
Label6.Caption = "2��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 2���� �Ͽ����ϴ�"
End If
If Label8.Caption = "1��" Then
Label6.Caption = "2��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 2���� �Ͽ����ϴ�"
End If
If Label3.Caption = "2��" Then
Label6.Caption = "3��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 3���� �Ͽ����ϴ�"
End If
If Label4.Caption = "2��" Then
Label6.Caption = "3��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 3���� �Ͽ����ϴ�"
End If
If Label5.Caption = "2��" Then
Label6.Caption = "3��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 3���� �Ͽ����ϴ�"
End If
If Label7.Caption = "2��" Then
Label6.Caption = "3��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 3���� �Ͽ����ϴ�"
End If
If Label8.Caption = "2��" Then
Label6.Caption = "3��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 3���� �Ͽ����ϴ�"
End If
If Label3.Caption = "3��" Then
Label6.Caption = "4��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 4���� �Ͽ����ϴ�"
End If
If Label4.Caption = "3��" Then
Label6.Caption = "4��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 4���� �Ͽ����ϴ�"
End If
If Label5.Caption = "3��" Then
Label6.Caption = "4��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 4���� �Ͽ����ϴ�"
End If
If Label7.Caption = "3��" Then
Label6.Caption = "4��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 4���� �Ͽ����ϴ�"
End If
If Label8.Caption = "3��" Then
Label6.Caption = "4��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 4���� �Ͽ����ϴ�"
End If
If Label3.Caption = "4��" Then
Label6.Caption = "5��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 5���� �Ͽ����ϴ�"
End If
If Label4.Caption = "4��" Then
Label6.Caption = "5��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 5���� �Ͽ����ϴ�"
End If
If Label5.Caption = "4��" Then
Label6.Caption = "5��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 5���� �Ͽ����ϴ�"
End If
If Label7.Caption = "4��" Then
Label6.Caption = "5��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 5���� �Ͽ����ϴ�"
End If
If Label8.Caption = "4��" Then
Label6.Caption = "5��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 5���� �Ͽ����ϴ�"
End If
If Label3.Caption = "5��" Then
Label6.Caption = "6��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 6���� �Ͽ����ϴ�"
End If
If Label4.Caption = "5��" Then
Label6.Caption = "6��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 6���� �Ͽ����ϴ�"
End If
If Label5.Caption = "5��" Then
Label6.Caption = "6��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 6���� �Ͽ����ϴ�"
End If
If Label7.Caption = "5��" Then
Label6.Caption = "6��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 6���� �Ͽ����ϴ�"
End If
If Label8.Caption = "5��" Then
Label6.Caption = "6��"
Label14.Caption = cnt & "��"
Label9.Caption = "�Ḹ���� 6���� �Ͽ����ϴ�"
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
Label7.Caption = "1��"
Label9.Caption = "������ 1���� �Ͽ����ϴ�"
Label15.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = "����"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = "����"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = "����"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If

End If
End If
End If
End If
End If
If Label7.Caption = "1��" Then
Label7.Caption = "2��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 2���� �Ͽ����ϴ�"
End If
If Label4.Caption = "1��" Then
Label7.Caption = "2��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 2���� �Ͽ����ϴ�"
End If
If Label5.Caption = "1��" Then
Label7.Caption = "2��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 2���� �Ͽ����ϴ�"
End If
If Label6.Caption = "1��" Then
Label7.Caption = "2��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 2���� �Ͽ����ϴ�"
End If
If Label8.Caption = "1��" Then
Label7.Caption = "2��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 2���� �Ͽ����ϴ�"
End If
If Label3.Caption = "2��" Then
Label7.Caption = "3��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 3���� �Ͽ����ϴ�"
End If
If Label4.Caption = "2��" Then
Label7.Caption = "3��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 3���� �Ͽ����ϴ�"
End If
If Label5.Caption = "2��" Then
Label7.Caption = "3��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 3���� �Ͽ����ϴ�"
End If
If Label6.Caption = "2��" Then
Label7.Caption = "3��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 3���� �Ͽ����ϴ�"
End If
If Label8.Caption = "2��" Then
Label7.Caption = "3��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 3���� �Ͽ����ϴ�"
End If
If Label3.Caption = "3��" Then
Label7.Caption = "4��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 4���� �Ͽ����ϴ�"
End If
If Label4.Caption = "3��" Then
Label7.Caption = "4��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 4���� �Ͽ����ϴ�"
End If
If Label5.Caption = "3��" Then
Label7.Caption = "4��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 4���� �Ͽ����ϴ�"
End If
If Label6.Caption = "3��" Then
Label7.Caption = "4��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 4���� �Ͽ����ϴ�"
End If
If Label8.Caption = "3��" Then
Label7.Caption = "4��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 4���� �Ͽ����ϴ�"
End If
If Label3.Caption = "4��" Then
Label7.Caption = "5��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 5���� �Ͽ����ϴ�"
End If
If Label4.Caption = "4��" Then
Label7.Caption = "5��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 5���� �Ͽ����ϴ�"
End If
If Label5.Caption = "4��" Then
Label7.Caption = "5��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 5���� �Ͽ����ϴ�"
End If
If Label6.Caption = "4��" Then
Label7.Caption = "5��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 5���� �Ͽ����ϴ�"
End If
If Label8.Caption = "4��" Then
Label7.Caption = "5��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 5���� �Ͽ����ϴ�"
End If
If Label3.Caption = "5��" Then
Label7.Caption = "6��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 6���� �Ͽ����ϴ�"
End If
If Label4.Caption = "5��" Then
Label7.Caption = "6��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 6���� �Ͽ����ϴ�"
End If
If Label5.Caption = "5��" Then
Label7.Caption = "6��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 6���� �Ͽ����ϴ�"
End If
If Label6.Caption = "5��" Then
Label7.Caption = "6��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 6���� �Ͽ����ϴ�"
End If
If Label8.Caption = "5��" Then
Label7.Caption = "6��"
Label15.Caption = cnt & "��"
Label9.Caption = "������ 6���� �Ͽ����ϴ�"
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
Label8.Caption = "1��"
Label9.Caption = "" & nam & "�� 1���� �Ͽ����ϴ�"
Label16.Caption = cnt & "��"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "��"
Label21.Visible = True
Label24.Caption = nam

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "��"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "��"
Label22.Visible = True
Label25.Caption = nam

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "��"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "��"
Label23.Visible = True
Label26.Caption = nam

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "��"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1��" Then
Label8.Caption = "2��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 2���� �Ͽ����ϴ�"
End If
If Label4.Caption = "1��" Then
Label8.Caption = "2��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 2���� �Ͽ����ϴ�"
End If
If Label5.Caption = "1��" Then
Label8.Caption = "2��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 2���� �Ͽ����ϴ�"
End If
If Label6.Caption = "1��" Then
Label8.Caption = "2��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 2���� �Ͽ����ϴ�"
End If
If Label7.Caption = "1��" Then
Label8.Caption = "2��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 2���� �Ͽ����ϴ�"
End If
If Label3.Caption = "2��" Then
Label8.Caption = "3��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 3���� �Ͽ����ϴ�"
End If
If Label4.Caption = "2��" Then
Label8.Caption = "3��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 3���� �Ͽ����ϴ�"
End If
If Label5.Caption = "2��" Then
Label8.Caption = "3��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 3���� �Ͽ����ϴ�"
End If
If Label6.Caption = "2��" Then
Label8.Caption = "3��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 3���� �Ͽ����ϴ�"
End If
If Label7.Caption = "2��" Then
Label8.Caption = "3��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 3���� �Ͽ����ϴ�"
End If
If Label3.Caption = "3��" Then
Label8.Caption = "4��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 4���� �Ͽ����ϴ�"
End If
If Label4.Caption = "3��" Then
Label8.Caption = "4��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 4���� �Ͽ����ϴ�"
End If
If Label5.Caption = "3��" Then
Label8.Caption = "4��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 4���� �Ͽ����ϴ�"
End If
If Label6.Caption = "3��" Then
Label8.Caption = "4��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 4���� �Ͽ����ϴ�"
End If
If Label7.Caption = "3��" Then
Label8.Caption = "4��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 4���� �Ͽ����ϴ�"
End If
If Label3.Caption = "4��" Then
Label8.Caption = "5��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 5���� �Ͽ����ϴ�"
End If
If Label4.Caption = "4��" Then
Label8.Caption = "5��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 5���� �Ͽ����ϴ�"
End If
If Label5.Caption = "4��" Then
Label8.Caption = "5��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 5���� �Ͽ����ϴ�"
End If
If Label6.Caption = "4��" Then
Label8.Caption = "5��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 5���� �Ͽ����ϴ�"
End If
If Label7.Caption = "4��" Then
Label8.Caption = "5��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 5���� �Ͽ����ϴ�"
End If
If Label3.Caption = "5��" Then
Label8.Caption = "6��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 6���� �Ͽ����ϴ�"
End If
If Label4.Caption = "5��" Then
Label8.Caption = "6��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 6���� �Ͽ����ϴ�"
End If
If Label5.Caption = "5��" Then
Label8.Caption = "6��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 6���� �Ͽ����ϴ�"
End If
If Label6.Caption = "5��" Then
Label8.Caption = "6��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 6���� �Ͽ����ϴ�"
End If
If Label7.Caption = "5��" Then
Label8.Caption = "6��"
Label16.Caption = cnt & "��"
Label9.Caption = "" & nam & "�� 6���� �Ͽ����ϴ�"
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
Label9.Caption = "���� �̻��ؾ��� ���η� �޸��� �ֽ��ϴ�"
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
Label9.Caption = "���� ���̸��� ���η� �޸��� �ֽ��ϴ�"
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
Label9.Caption = "���� ���αⰡ ���η� �޸��� �ֽ��ϴ�"
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
Label9.Caption = "���� �Ḹ���� ���η� �޸��� �ֽ��ϴ�"
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
Label9.Caption = "���� ������ ���η� �޸��� �ֽ��ϴ�"
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
Label9.Caption = "���� " & nam & "�� ���η� �޸��� �ֽ��ϴ�"
End Select
End If
End If
End If
End If
End If
If Image6.Left <= -200 Then
MsgBox "����� ����忡�� ������߽��ϴ�" + Chr(13) + Chr(13) _
        + "�׷��� ������ ��⸦ ó������ �����մϴ�."
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
