VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "�ܾ� ���"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "Ÿ�� ���� V1.0 8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   7575
   ScaleWidth      =   8310
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����"
      Height          =   615
      Left            =   6840
      Style           =   1  '�׷���
      TabIndex        =   26
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Timer Timer33 
      Interval        =   1000
      Left            =   7920
      Top             =   3600
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      TabIndex        =   24
      Text            =   "100"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Timer Timer32 
      Interval        =   100
      Left            =   7920
      Top             =   3240
   End
   Begin VB.Timer Timer31 
      Interval        =   100
      Left            =   7920
      Top             =   2880
   End
   Begin VB.Timer Timer30 
      Interval        =   100
      Left            =   7920
      Top             =   2520
   End
   Begin VB.Timer Timer29 
      Interval        =   100
      Left            =   7920
      Top             =   2160
   End
   Begin VB.Timer Timer28 
      Left            =   7920
      Top             =   1800
   End
   Begin VB.Timer Timer27 
      Interval        =   100
      Left            =   7920
      Top             =   1440
   End
   Begin VB.Timer Timer26 
      Interval        =   100
      Left            =   7920
      Top             =   1080
   End
   Begin VB.Timer Timer25 
      Interval        =   100
      Left            =   7920
      Top             =   720
   End
   Begin VB.Timer Timer24 
      Interval        =   100
      Left            =   7920
      Top             =   360
   End
   Begin VB.Timer Timer23 
      Interval        =   100
      Left            =   7920
      Top             =   0
   End
   Begin VB.Timer Timer22 
      Interval        =   100
      Left            =   7560
      Top             =   0
   End
   Begin VB.Timer Timer21 
      Interval        =   100
      Left            =   7200
      Top             =   0
   End
   Begin VB.Timer Timer20 
      Interval        =   100
      Left            =   6840
      Top             =   0
   End
   Begin VB.Timer Timer19 
      Interval        =   100
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer Timer18 
      Interval        =   100
      Left            =   6120
      Top             =   0
   End
   Begin VB.Timer Timer17 
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer16 
      Interval        =   100
      Left            =   5400
      Top             =   0
   End
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin VB.Timer Timer14 
      Interval        =   100
      Left            =   4680
      Top             =   0
   End
   Begin VB.Timer Timer13 
      Interval        =   100
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�޴�"
      Height          =   615
      Left            =   5280
      Style           =   1  '�׷���
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ó������"
      Height          =   615
      Left            =   3600
      Style           =   1  '�׷���
      TabIndex        =   18
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�ߴ��ϱ�"
      Height          =   615
      Left            =   1920
      Style           =   1  '�׷���
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�����ϱ�(&S)"
      Height          =   615
      Left            =   240
      Style           =   1  '�׷���
      TabIndex        =   16
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "���� �ð� :"
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
      TabIndex        =   23
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
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
      Left            =   6360
      TabIndex        =   22
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "���� ���� :"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   9360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   6120
      Shape           =   2  'Ÿ����
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   4200
      Shape           =   2  'Ÿ����
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   2280
      Shape           =   2  'Ÿ����
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   360
      Shape           =   2  'Ÿ����
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   6120
      Shape           =   2  'Ÿ����
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   4200
      Shape           =   2  'Ÿ����
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   2280
      Shape           =   2  'Ÿ����
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   360
      Shape           =   2  'Ÿ����
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   6120
      Shape           =   2  'Ÿ����
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   4200
      Shape           =   2  'Ÿ����
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   2280
      Shape           =   2  'Ÿ����
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   360
      Shape           =   2  'Ÿ����
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   6120
      Shape           =   2  'Ÿ����
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   4200
      Shape           =   2  'Ÿ����
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   2280
      Shape           =   2  'Ÿ����
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   735
      Left            =   360
      Shape           =   2  'Ÿ����
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   6360
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   4440
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   600
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   6360
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   4440
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   6360
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   6360
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
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
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ����(200), b, bb, c, d, e, f, g, h, i, k, l, m, n, o, p, q, r, cc, dd, ee, ff, gg, hh, ii, kk, ll, mm, nn, oo, pp, qq, rr, aaa, ab

Private Sub Command1_Click()
Text1.IMEMode = vbIMEModeHangul '�ѱ�
aaa = Val(Text2.Text)
Timer1.Enabled = True
Timer1.Enabled = True
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
Timer33.Enabled = True
For i = 0 To 15
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Label1(i).Visible = False
Next
Timer1.Interval = Int(Rnd(100) * 1000)
Timer2.Interval = Int(Rnd(100) * 1000)
Timer3.Interval = Int(Rnd(100) * 1000)
Timer4.Interval = Int(Rnd(100) * 1000)
Timer5.Interval = Int(Rnd(100) * 1000)
Timer6.Interval = Int(Rnd(100) * 1000)
Timer7.Interval = Int(Rnd(100) * 1000)
Timer8.Interval = Int(Rnd(100) * 1000)
Timer9.Interval = Int(Rnd(100) * 1000)
Timer10.Interval = Int(Rnd(100) * 1500)
Timer11.Interval = Int(Rnd(100) * 1500)
Timer12.Interval = Int(Rnd(100) * 1500)
Timer13.Interval = Int(Rnd(100) * 1500)
Timer14.Interval = Int(Rnd(100) * 1500)
Timer15.Interval = Int(Rnd(100) * 1500)
Timer16.Interval = Int(Rnd(100) * 1500)
Timer17.Interval = Int(Rnd(100) * 2000)
Timer18.Interval = Int(Rnd(100) * 2000)
Timer19.Interval = Int(Rnd(100) * 2000)
Timer20.Interval = Int(Rnd(100) * 2000)
Timer21.Interval = Int(Rnd(100) * 2000)
Timer22.Interval = Int(Rnd(100) * 2000)
Timer23.Interval = Int(Rnd(100) * 2500)
Timer24.Interval = Int(Rnd(100) * 2500)
Timer25.Interval = Int(Rnd(100) * 2500)
Timer26.Interval = Int(Rnd(100) * 2500)
Timer27.Interval = Int(Rnd(100) * 2500)
Timer28.Interval = Int(Rnd(100) * 2500)
Timer29.Interval = Int(Rnd(100) * 3000)
Timer30.Interval = Int(Rnd(100) * 3000)
Timer31.Interval = Int(Rnd(100) * 3000)
Timer32.Interval = Int(Rnd(100) * 3000)
Label5.Visible = True
Label5.Enabled = True
Text2.Enabled = False
Text2.Visible = False
Text1.SetFocus
Text1.Text = ""
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True

End Sub

Private Sub Command3_Click()
aaa = Val(Text2.Text)
Label3.Caption = 0
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
For i = 0 To 15
Label1(i).Visible = False
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)
Next
Label1(0).Top = 840
Label1(1).Top = 840
Label1(2).Top = 840
Label1(3).Top = 840
Label1(4).Top = 2160
Label1(5).Top = 2160
Label1(6).Top = 2160
Label1(7).Top = 2160
Label1(8).Top = 3480
Label1(9).Top = 3480
Label1(10).Top = 3480
Label1(11).Top = 3480
Label1(12).Top = 4800
Label1(13).Top = 4800
Label1(14).Top = 4800
Label1(15).Top = 4800
Label5.Visible = False
Label5.Enabled = False
Text2.Visible = True
Text2.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
MsgBox "<�ܾ���� ����>" + Chr(13) + Chr(13) _
        + "�ȳ��ϼ���." + Chr(13) + Chr(13) _
        + "�ܾ����� �δ��� ���� ������ �����Ѱ̴ϴ�" + Chr(13) + Chr(13) _
        + "�ܾ ������ �� ���ð� ���ֽñ� �ٶ��ϴ�. ����ְ� �ϼ���.^^"
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
Label3.Caption = 0
Label5.Enabled = False
Label5.Visible = False
Text2.Enabled = True
Text2.Visible = True
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
    For i = 0 To 15
        If Trim(Text1.Text) = Label1(i).Caption Then
            If Text1.Text <> "" Then
            Label3.Caption = Label3.Caption + 1
            Label1(i).Caption = ""
            End If
         End If
    Next
    Text1.Text = ""
ab = Label3.Caption

End If
End Sub

Private Sub Timer1_Timer()
b = Int(Rnd(10) * 100)
Label1(0).Top = Label1(0).Top - b
If Label1(0).Top <= 140 Then
Label1(0).Visible = True
End If
If Label1(0).Top <= 90 Then
b = 0
Label1(0).Top = 90
Timer1.Enabled = False
Timer17.Enabled = True
End If
End Sub

Private Sub Timer10_Timer()
l = Int(Rnd(10) * 100)
Label1(9).Top = Label1(9).Top - l
If Label1(9).Top <= 2760 Then
Label1(9).Visible = True
End If
If Label1(9).Top <= 2640 Then
l = 0
Label1(9).Top = 2640
Timer10.Enabled = False
Timer26.Enabled = True
End If

End Sub

Private Sub Timer11_Timer()
m = Int(Rnd(10) * 100)
Label1(10).Top = Label1(10).Top - m
If Label1(10).Top <= 2760 Then
Label1(10).Visible = True
End If
If Label1(10).Top <= 2640 Then
m = 0
Label1(10).Top = 2640
Timer11.Enabled = False
Timer27.Enabled = True
End If

End Sub

Private Sub Timer12_Timer()
n = Int(Rnd(10) * 100)
Label1(11).Top = Label1(11).Top - n
If Label1(11).Top <= 2760 Then
Label1(11).Visible = True
End If
If Label1(11).Top <= 2640 Then
n = 0
Label1(11).Top = 2640
Timer12.Enabled = False
Timer28.Enabled = True
End If

End Sub

Private Sub Timer13_Timer()
o = Int(Rnd(10) * 100)
Label1(12).Top = Label1(12).Top - o
If Label1(12).Top <= 4080 Then
Label1(12).Visible = True
End If
If Label1(12).Top <= 3960 Then
o = 0
Label1(12).Top = 3960
Timer13.Enabled = False
Timer29.Enabled = True
End If

End Sub

Private Sub Timer14_Timer()
p = Int(Rnd(10) * 100)
Label1(13).Top = Label1(13).Top - p
If Label1(13).Top <= 4080 Then
Label1(13).Visible = True
End If
If Label1(13).Top <= 3960 Then
p = 0
Label1(13).Top = 3960
Timer14.Enabled = False
Timer30.Enabled = True
End If

End Sub

Private Sub Timer15_Timer()
q = Int(Rnd(10) * 100)
Label1(14).Top = Label1(14).Top - q
If Label1(14).Top <= 4080 Then
Label1(14).Visible = True
End If
If Label1(14).Top <= 3960 Then
q = 0
Label1(14).Top = 3960
Timer15.Enabled = False
Timer31.Enabled = True
End If

End Sub

Private Sub Timer16_Timer()
r = Int(Rnd(10) * 100)
Label1(15).Top = Label1(15).Top - r
If Label1(15).Top <= 4080 Then
Label1(15).Visible = True
End If
If Label1(15).Top <= 3960 Then
r = 0
Label1(15).Top = 3960
Timer16.Enabled = False
Timer32.Enabled = True
End If

End Sub

Private Sub Timer17_Timer()
bb = Int(Rnd(10) * 100)
Label1(0).Top = Label1(0).Top + bb
If Label1(0).Top >= 140 Then
Label1(0).Visible = False
End If
If Label1(0).Top >= 840 Then
Timer17.Enabled = False
Timer1.Enabled = True
Label1(0).Caption = ""
a = Int(Rnd(1) * 200)
Label1(0).Caption = ����(a)
End If
End Sub

Private Sub Timer18_Timer()
cc = Int(Rnd(10) * 100)
Label1(1).Top = Label1(1).Top + cc
If Label1(1).Top >= 140 Then
Label1(1).Visible = False
End If
If Label1(1).Top >= 840 Then
Timer18.Enabled = False
Timer2.Enabled = True
Label1(1).Caption = ""
a = Int(Rnd(1) * 200)
Label1(1).Caption = ����(a)
End If

End Sub

Private Sub Timer19_Timer()
dd = Int(Rnd(10) * 100)
Label1(2).Top = Label1(2).Top + dd
If Label1(2).Top >= 140 Then
Label1(2).Visible = False
End If
If Label1(2).Top >= 840 Then
Timer19.Enabled = False
Timer3.Enabled = True
Label1(2).Caption = ""
a = Int(Rnd(2) * 200)
Label1(2).Caption = ����(a)
End If

End Sub

Private Sub Timer2_Timer()
c = Int(Rnd(10) * 100)
Label1(1).Top = Label1(1).Top - c
If Label1(1).Top <= 140 Then
Label1(1).Visible = True
End If
If Label1(1).Top <= 90 Then
c = 0
Label1(1).Top = 90
Timer2.Enabled = False
Timer18.Enabled = True
End If
End Sub

Private Sub Timer20_Timer()
ee = Int(Rnd(10) * 100)
Label1(3).Top = Label1(3).Top + ee
If Label1(3).Top >= 140 Then
Label1(3).Visible = False
End If
If Label1(3).Top >= 840 Then
Timer20.Enabled = False
Timer4.Enabled = True
Label1(3).Caption = ""
a = Int(Rnd(1) * 200)
Label1(3).Caption = ����(a)
End If

End Sub

Private Sub Timer21_Timer()
ff = Int(Rnd(10) * 100)
Label1(4).Top = Label1(4).Top + ff
If Label1(4).Top >= 1440 Then
Label1(4).Visible = False
End If
If Label1(4).Top >= 2160 Then
Timer21.Enabled = False
Timer5.Enabled = True
Label1(4).Caption = ""
a = Int(Rnd(1) * 200)
Label1(4).Caption = ����(a)
End If

End Sub

Private Sub Timer22_Timer()
gg = Int(Rnd(10) * 100)
Label1(5).Top = Label1(5).Top + gg
If Label1(5).Top >= 1440 Then
Label1(5).Visible = False
End If
If Label1(5).Top >= 2160 Then
Timer22.Enabled = False
Timer6.Enabled = True
Label1(5).Caption = ""
a = Int(Rnd(1) * 200)
Label1(5).Caption = ����(a)
End If

End Sub
    
Private Sub Timer23_Timer()
hh = Int(Rnd(10) * 100)
Label1(6).Top = Label1(6).Top + hh
If Label1(6).Top >= 1440 Then
Label1(6).Visible = False
End If
If Label1(6).Top >= 2160 Then
Timer23.Enabled = False
Timer7.Enabled = True
Label1(6).Caption = ""
a = Int(Rnd(1) * 200)
Label1(6).Caption = ����(a)
End If
End Sub

Private Sub Timer24_Timer()
ii = Int(Rnd(10) * 100)
Label1(7).Top = Label1(7).Top + ii
If Label1(7).Top >= 1440 Then
Label1(7).Visible = False
End If
If Label1(7).Top >= 2160 Then
Timer24.Enabled = False
Timer8.Enabled = True
Label1(7).Caption = ""
a = Int(Rnd(1) * 200)
Label1(7).Caption = ����(a)
End If

End Sub

Private Sub Timer25_Timer()
kk = Int(Rnd(10) * 100)
Label1(8).Top = Label1(8).Top + kk
If Label1(8).Top >= 2760 Then
Label1(8).Visible = False
End If
If Label1(8).Top >= 3480 Then
Timer25.Enabled = False
Timer9.Enabled = True
Label1(8).Caption = ""
a = Int(Rnd(1) * 200)
Label1(8).Caption = ����(a)
End If

End Sub

Private Sub Timer26_Timer()
ll = Int(Rnd(10) * 100)
Label1(9).Top = Label1(9).Top + ll
If Label1(9).Top >= 2760 Then
Label1(9).Visible = False
End If
If Label1(9).Top >= 3480 Then
Timer26.Enabled = False
Timer10.Enabled = True
Label1(9).Caption = ""
a = Int(Rnd(1) * 200)
Label1(9).Caption = ����(a)
End If

End Sub

Private Sub Timer27_Timer()
mm = Int(Rnd(10) * 100)
Label1(10).Top = Label1(10).Top + mm
If Label1(10).Top >= 2760 Then
Label1(10).Visible = False
End If
If Label1(10).Top >= 3480 Then
Timer27.Enabled = False
Timer11.Enabled = True
Label1(10).Caption = ""
a = Int(Rnd(1) * 200)
Label1(10).Caption = ����(a)
End If
End Sub

Private Sub Timer28_Timer()
nn = Int(Rnd(10) * 100)
Label1(11).Top = Label1(11).Top + nn
If Label1(11).Top >= 2760 Then
Label1(11).Visible = False
End If
If Label1(11).Top >= 3480 Then
Timer28.Enabled = False
Timer12.Enabled = True
Label1(11).Caption = ""
a = Int(Rnd(1) * 200)
Label1(11).Caption = ����(a)
End If

End Sub

Private Sub Timer29_Timer()
oo = Int(Rnd(10) * 100)
Label1(12).Top = Label1(12).Top + oo
If Label1(12).Top >= 4080 Then
Label1(12).Visible = False
End If
If Label1(12).Top >= 4800 Then
Timer29.Enabled = False
Timer13.Enabled = True
Label1(12).Caption = ""
a = Int(Rnd(1) * 200)
Label1(12).Caption = ����(a)
End If

End Sub

Private Sub Timer3_Timer()
d = Int(Rnd(10) * 100)
Label1(2).Top = Label1(2).Top - d
If Label1(2).Top <= 140 Then
Label1(2).Visible = True
End If
If Label1(2).Top <= 90 Then
d = 0
Label1(2).Top = 90
Timer3.Enabled = False
Timer19.Enabled = True
End If

End Sub

Private Sub Timer30_Timer()
pp = Int(Rnd(10) * 100)
Label1(13).Top = Label1(13).Top + pp
If Label1(13).Top >= 4080 Then
Label1(13).Visible = False
End If
If Label1(13).Top >= 4800 Then
Timer30.Enabled = False
Timer14.Enabled = True
Label1(13).Caption = ""
a = Int(Rnd(1) * 200)
Label1(13).Caption = ����(a)
End If


End Sub

Private Sub Timer31_Timer()
qq = Int(Rnd(10) * 100)
Label1(14).Top = Label1(14).Top + qq
If Label1(14).Top >= 4080 Then
Label1(14).Visible = False
End If
If Label1(14).Top >= 4800 Then
Timer31.Enabled = False
Timer15.Enabled = True
Label1(14).Caption = ""
a = Int(Rnd(1) * 200)
Label1(14).Caption = ����(a)
End If


End Sub

Private Sub Timer32_Timer()
rr = Int(Rnd(10) * 100)
Label1(15).Top = Label1(15).Top + rr
If Label1(15).Top >= 4080 Then
Label1(15).Visible = False
End If
If Label1(15).Top >= 4800 Then
Timer32.Enabled = False
Timer16.Enabled = True
Label1(15).Caption = ""
a = Int(Rnd(1) * 200)
Label1(15).Caption = ����(a)
End If


End Sub

Private Sub Timer33_Timer()
aaa = aaa - 1
Label5.Caption = aaa
If Label5 = 0 Then
MsgBox "�ð��� �� �Ǿ����ϴ�."
MsgBox "����� ���� ������ " & ab & "�� �Դϴ�."
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
For i = 0 To 15
Label1(i).Visible = False
a = Int(Rnd(1) * 200)
Label1(i).Caption = ����(a)

Next
Label1(0).Top = 840
Label1(1).Top = 840
Label1(2).Top = 840
Label1(3).Top = 840
Label1(4).Top = 2160
Label1(5).Top = 2160
Label1(6).Top = 2160
Label1(7).Top = 2160
Label1(8).Top = 3480
Label1(9).Top = 3480
Label1(10).Top = 3480
Label1(11).Top = 3480
Label1(12).Top = 4800
Label1(13).Top = 4800
Label1(14).Top = 4800
Label1(15).Top = 4800
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Label5.Visible = False
Label5.Enabled = False
Text2.Enabled = True
Text2.Visible = True
Text1.Text = ""
Label3.Caption = 0
End If
End Sub

Private Sub Timer4_Timer()
e = Int(Rnd(10) * 100)
Label1(3).Top = Label1(3).Top - e
If Label1(3).Top <= 140 Then
Label1(3).Visible = True
End If
If Label1(3).Top <= 90 Then
e = 0
Label1(3).Top = 90
Timer4.Enabled = False
Timer20.Enabled = True
End If

End Sub

Private Sub Timer5_Timer()
f = Int(Rnd(10) * 100)
Label1(4).Top = Label1(4).Top - f
If Label1(4).Top <= 1440 Then
Label1(4).Visible = True
End If
If Label1(4).Top <= 1320 Then
f = 0
Label1(4).Top = 1320
Timer5.Enabled = False
Timer21.Enabled = True
End If

End Sub

Private Sub Timer6_Timer()
g = Int(Rnd(10) * 100)
Label1(5).Top = Label1(5).Top - g
If Label1(5).Top <= 1440 Then
Label1(5).Visible = True
End If
If Label1(5).Top <= 1320 Then
g = 0
Label1(5).Top = 1320
Timer6.Enabled = False
Timer22.Enabled = True
End If

End Sub

Private Sub Timer7_Timer()
h = Int(Rnd(10) * 100)
Label1(6).Top = Label1(6).Top - h
If Label1(6).Top <= 1440 Then
Label1(6).Visible = True
End If
If Label1(6).Top <= 1320 Then
h = 0
Label1(6).Top = 1320
Timer7.Enabled = False
Timer23.Enabled = True
End If

End Sub

Private Sub Timer8_Timer()
i = Int(Rnd(10) * 100)
Label1(7).Top = Label1(7).Top - i
If Label1(7).Top <= 1440 Then
Label1(7).Visible = True
End If
If Label1(7).Top <= 1320 Then
i = 0
Label1(7).Top = 1320
Timer8.Enabled = False
Timer24.Enabled = True
End If
End Sub

Private Sub Timer9_Timer()
k = Int(Rnd(10) * 100)
Label1(8).Top = Label1(8).Top - k
If Label1(8).Top <= 2760 Then
Label1(8).Visible = True
End If
If Label1(8).Top <= 2640 Then
k = 0
Label1(8).Top = 2640
Timer9.Enabled = False
Timer25.Enabled = True
End If

End Sub
