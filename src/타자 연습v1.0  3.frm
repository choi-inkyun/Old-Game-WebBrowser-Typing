VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "ª����"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ÿ�� ����v1.0  3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4755
   ScaleWidth      =   6315
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "�޴�(&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�ߴ��ϱ�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  '�׷���
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
      Caption         =   "�����ϱ�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
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
      Caption         =   "�ִ�Ÿ��:"
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
      Caption         =   "Ÿ��:"
      Height          =   495
      Left            =   240
      TabIndex        =   36
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "Ȯ�� :"
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
      Left            =   1800
      TabIndex        =   32
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Ȯ�� :"
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
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   30
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "���᰹��:"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
Dim ����(100), v, k, c, d
Dim cnt As Byte

Private Sub Command1_Click()
Text1.IMEMode = vbIMEModeHangul '�ѱ�

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
����(1) = "�� �Ѹ���� õ�ɺ��� ���´�."
����(2) = "������������� �㸻���㰡��´�."
����(3) = "���¸��� ���;� ���� ���� ����."
����(4) = "�� �Ϳ� �� �б�"
����(5) = "������ �⿪�ڵ� �𸥴�."
����(6) = "����� �̿��� �� ģô���� ����."
����(7) = "���� ���� �峯�̴�."
����(8) = "���� ��˸��ϴ�."
����(9) = "����� �� ��"
����(10) = "�������� ������ �ٶ��� ���� ����."
����(11) = "���� �⺰�� ���� �ʴ´�."
����(12) = "���信 ���丮"
����(13) = "���� �⺰�� ���� �ʴ´�."
����(14) = "��õ���� �� ����."
����(15) = "���� �Ӹ� �� �Ѹ� �ǵ���"
����(16) = "���� ž�� ��������"
����(17) = "��� ���� �� ��������."
����(18) = "�� ��� ��"
����(19) = "�� �԰� �� �Դ´�."
����(20) = "���� ���� ���� ���δ�."
����(21) = "������ �� �Ա�"
����(22) = "�ü� �԰� �̾��ñ�"
����(23) = "���ٸ��� �ε�� ���� �ǳʶ�"
����(24) = "�ް��� ���� ġ��"
����(25) = "������ ���̿� �ȵ����� ���̴�."
����(26) = "�� ¤�� ���ġ��"
����(27) = "�Ƿ� �ְ� ���� �޴´�."
����(28) = "�ٴ� ���µ� ���� ����."
����(29) = "�� �Ұ� �ܾ簣 ��ģ��."
����(30) = "���� �� �q��"
����(31) = "�ƴ� �Ǹ� ���� ſ"
����(32) = "�� ���� ���� õ�� ����."
����(33) = "���� ���� ������"
����(34) = "�����浵 �� �������� ���۵ȴ�."
����(35) = "���� �ϴÿ� ������"
����(36) = "�ƴ� ���� ��"
����(37) = "�����¿ܳ��� �ٸ����� ������."
����(38) = "��� ���ڸԱ�"
����(39) = "�������� �� ���� ��"
����(40) = "�����̵� ������ ��Ʋ�Ѵ�."
����(41) = "���࿡ ���� ������"
����(42) = "�Ǵ� ������ ���ϴ�."
����(43) = "�������� �� �ߵ��� ��´�."
����(44) = "�� �ҿ� �Ѿ��."
����(45) = "���� �ϵ� �ϻ� �Ҽ�"
����(46) = "������ ���� �ζѸ��� ������."
����(47) = "�Լ��� ħ�̳� �ٸ���"
����(48) = "���� ���� ħ ������"
����(49) = "�칰�� �ĵ� �� �칰�� �Ķ�"
����(50) = "�����̸��ƾ� �Ʒ����� ����"
����(51) = "������ �ϰ� �ο��� ��������"
����(52) = "�� �� �信 ��θ���?"
����(53) = "�Ϸ������� �� ������ �� �𸥴�"
����(54) = "�ϴ��� �������� �ھƳ� ������ �ִ�"
����(55) = "�ϴ��� ������"
����(56) = "�ΰ� ���� ������ ����"
����(57) = "���� ���̶� �Ͽ��� ���� ��´�"
����(58) = "���� ������ ���´�"
����(59) = "Ƽ�� ��� �»�"
����(60) = "ū �Ͽ��� ū �Ҹ�����"
����(61) = "ū ���׵� ���̱������� ��������"
����(62) = "Į�� �� ����"
����(63) = "õ�� �浵 �� ��������"
����(64) = "������ ���� �ٰ����� �鿡 ���� ����"
����(65) = "�����̵� ������ ��Ʋ�Ѵ�"
����(66) = "�㱸���� ã�´�"
����(67) = "�㱸�ۿ��� �� �� ���� �ִ�"
����(68) = "���� ���� �Կ� ����"
����(69) = "���ο��� �� �°� �Ѱ� ���� �� ����"
����(70) = "����� �۾Ƶ� ������ ����"
����(71) = "�� ���� �� �ٱ�"
����(72) = "�� ���� �����"
����(73) = "���� ���߰� �� �ʴ�"
����(74) = "�ڶ� ������ ���ٺ��� �˾ƺ���"
����(75) = "�ڶ󿡰� ��� ���� �ܶѰ� ���� �����"
����(76) = "�ڴ� �� ��ħ�ֱ�"
����(77) = "���� ���� ���� �д�"
����(78) = "�ƹ��� �ٺ��� �ٴ� �㸮 �ž� �� ����"
����(79) = "�ƴ� ���߿� ȫ�α�"
����(80) = "�� ���̸� ���굵 ���Ѵ�"
����(81) = "�ƴ� �浵 ���� ����"
����(82) = "�ƴ� �� ���ҿ� ���⳯��"
����(83) = "���� �� �Ա�"
����(84) = "������ ���̴�"
����(85) = "�÷翡 �� �ۺױ�"
����(86) = "��� �ٴϱ� �����̵� �ڴ�"
����(87) = "���� �� �q��"
����(88) = "�۰��� ������ ����"
����(89) = "�� �Ұ� �ܾ簣 ��ģ��"
����(90) = "�칰 �� ������"
����(91) = "��� ���ڸԱ�"
����(92) = "�Ұ��� ��õ ��������"
����(93) = "���� ������"
����(94) = "������ Ƽ�� �ִ�"
����(95) = "���� �Ʊ� ��� ã�´�"
����(96) = "��� ������ �öѱⰡ ��Ų��"
����(97) = "���� ���� ���� �ǳʶ�"
����(98) = "�� �� �׿� �� ������"
����(99) = "�޸� ��Ű�� ���� ��´�"
����(100) = "���� ���� ��Ӵ�"


a = Int(Rnd(1) * 100)
Label6.Caption = ����(a)
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Label6.Visible = False
k = Len(����(a))
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
           Label6.Caption = ����(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = ����(a)
           End If
           k = Len(����(a))
           For i = 0 To k
           Label1(i).Caption = "x"
           Next
           Label8.Caption = "�¾Ҵ�"
        ElseIf v <> Label6.Caption Then
         Label8.Caption = "Ʋ�ȴ�"
           a = Int(Rnd(1) * 100)
           Label6.Caption = ����(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = ����(a)
          End If
           k = Len(����(a))
           
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
