VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "����"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "����4.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6690
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "���� ����"
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
      Picture         =   "����4.frx":030A
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
      Picture         =   "����4.frx":2AAC
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
Label1.Caption = "�ȳ��ϼ���. Air Force �� ���� �Դϴ�." + Chr(13) + Chr(13) _
        + "Space = �̻���, Ctrl = �ʻ��, ������Ű�� �����̽ø� �˴ϴ�" + Chr(13) + Chr(13) _
        + "1, 2��° ���������� ���еǾ� �ֽ��ϴ�...�׷� ����ְ� ���ñ� �ٶ��ϴ�." + Chr(13) + Chr(13) _
        + "���׳� ��Ÿ �ǹ������������ø� dingpong@hitel.net ���� ���Ϻ�Ź�帳�ϴ�."
Label2.Caption = "                                              �ó�����" + Chr(13) + Chr(13) _
                + "���ΰ��� ��Ÿ�� ����Ų��....������ �����ִ� ������� ������ �ο��" + Chr(13) + Chr(13) _
                + "���� ���� �Ź��� ����⸦ �μ���...�ᱹ ���� Ÿ���ִ� �������� �μ��� �Ÿ�" + Chr(13) + Chr(13) _
                + "������ ����� �¸��ϰ� �Ǵ� ���̴� ������..Ȧ�� �ο�� ������ ���Ʒ� �¿쿡��" + Chr(13) + Chr(13) _
                + "������� ���� ������ ������� �̻����� ���� �ձ��� ���̴� ���� ����� �ƴϴ�." + Chr(13) + Chr(13) _
                + "���� �װͿ� �ڽ��� ����� 3�밡 ��� ���ߴ��Ѵٸ� ����� ���ԵǴ� ���̴�." + Chr(13) + Chr(13) _
                + "����� ���.."
Label3.Caption = ": �̰��� ������ �ʻ�⸦ �ٽ� �����ְԵȴ�."
Label4.Caption = ": ����ȭ�� �����ʾƷ� ��Ÿ���� ���̴�." + Chr(13) + Chr(13) _
                + "�̰��� ������ �ʻ�⸦ �����ִ�."
End Sub

