VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "�ڸ�����"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "Ÿ�� ���� V1.0 6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5145
   ScaleWidth      =   6585
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command4 
      Caption         =   "�����"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ѱۿ���"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����ϱ�"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�޴���"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "�ؽ�Ʈ�ڽ����� ġ�ø� �˴ϴ�..�׸��� ���� ���ڰ� ������ ������ �����ϱ⸦ �ѹ��� �����ּ���. "
      Height          =   975
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "�������� :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "������ �ڸ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim �ѱ��ڸ�(26), a, �����ڸ�(27)

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Label1.Visible = True
Label4.Caption = 0
End Sub

Private Sub Command3_Click()
Label1.Visible = False
�ѱ��ڸ�(1) = "��"
�ѱ��ڸ�(2) = "��"
�ѱ��ڸ�(3) = "��"
�ѱ��ڸ�(4) = "��"
�ѱ��ڸ�(5) = "��"
�ѱ��ڸ�(6) = "��"
�ѱ��ڸ�(7) = "��"
�ѱ��ڸ�(8) = "��"
�ѱ��ڸ�(9) = "��"
�ѱ��ڸ�(10) = "��"
�ѱ��ڸ�(11) = "��"
�ѱ��ڸ�(12) = "��"
�ѱ��ڸ�(13) = "��"
�ѱ��ڸ�(26) = "��"
�ѱ��ڸ�(15) = "��"
�ѱ��ڸ�(16) = "��"
�ѱ��ڸ�(17) = "��"
�ѱ��ڸ�(18) = "��"
�ѱ��ڸ�(19) = "��"
�ѱ��ڸ�(20) = "��"
�ѱ��ڸ�(21) = "��"
�ѱ��ڸ�(22) = "��"
�ѱ��ڸ�(23) = "��"
�ѱ��ڸ�(24) = "��"
�ѱ��ڸ�(25) = "��"
�ѱ��ڸ�(26) = "��"
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Command3.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command4_Click()
Label1.Visible = False
�����ڸ�(1) = "A"
�����ڸ�(2) = "B"
�����ڸ�(3) = "C"
�����ڸ�(4) = "D"
�����ڸ�(5) = "E"
�����ڸ�(6) = "F"
�����ڸ�(7) = "G"
�����ڸ�(8) = "H"
�����ڸ�(9) = "I"
�����ڸ�(10) = "J"
�����ڸ�(11) = "K"
�����ڸ�(12) = "L"
�����ڸ�(13) = "M"
�����ڸ�(14) = "N"
�����ڸ�(15) = "O"
�����ڸ�(16) = "P"
�����ڸ�(18) = "Q"
�����ڸ�(17) = "R"
�����ڸ�(19) = "S"
�����ڸ�(20) = "T"
�����ڸ�(21) = "U"
�����ڸ�(22) = "V"
�����ڸ�(23) = "W"
�����ڸ�(24) = "U"
�����ڸ�(25) = "X"
�����ڸ�(26) = "Y"
�����ڸ�(27) = "Z"
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Form_Load()
Randomize
Label1.Visible = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Command3.Enabled = False Then
If Label1.Caption = "��" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "��" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "��" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If




End If

If Command4.Enabled = False Then
If Label1.Caption = "A" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "B" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "C" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If


If Label1.Caption = "D" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "E" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "F" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "G" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "H" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "I" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "J" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "K" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "L" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "M" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "N" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "O" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "P" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Q" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "R" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "S" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "T" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "U" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "V" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "W" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "X" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Y" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "Z" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
End If
If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If

End Sub

