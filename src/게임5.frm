VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  '����
   Caption         =   "Form6"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form6"
   ScaleHeight     =   6840
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox Picture11 
      Appearance      =   0  '���
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5040
      Picture         =   "����5.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   4320
      Width           =   510
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  '���
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   1920
      Picture         =   "����5.frx":0C42
      ScaleHeight     =   930
      ScaleWidth      =   720
      TabIndex        =   13
      Top             =   4320
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ϱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   12
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "�̻��� ����"
      Height          =   3135
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      Begin VB.PictureBox Picture9 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   10
         Top             =   2160
         Width           =   735
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����� ����"
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      Begin VB.PictureBox Picture6 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   6
         Top             =   1800
         Width           =   855
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   705
         TabIndex        =   4
         Top             =   1680
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   705
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      Caption         =   "�����ϰ� ���� ���� Ŭ���� �ֽʽÿ�..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path & "\���ú�����2.BMP")
Picture2.Picture = LoadPicture(App.Path & "\���ú�����.BMP")
Picture3.Picture = LoadPicture(App.Path & "\�������_��.BMP")
Picture4.Picture = LoadPicture(App.Path & "\������.BMP")
Picture5.Picture = LoadPicture(App.Path & "\air��.BMP")
Picture6.Picture = LoadPicture(App.Path & "\Enemy��.BMP")
Picture7.Picture = LoadPicture(App.Path & "\���̻��Ϲ�.BMP")
Picture8.Picture = LoadPicture(App.Path & "\�̻��Ϲ�.BMP")
Picture9.Picture = LoadPicture(App.Path & "\shoot��.BMP")
End Sub

Private Sub Picture1_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\���ú������2.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\���ú�����2.BMP")
Picture10.Picture = LoadPicture(App.Path & "\���ú�����2.BMP")

End Sub

Private Sub Picture2_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\���ú������.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\���ú�����.BMP")
Picture10.Picture = LoadPicture(App.Path & "\���ú�����.BMP")

End Sub

Private Sub Picture3_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\�������_��.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\�������_��.BMP")
Picture10.Picture = LoadPicture(App.Path & "\�������_��.BMP")

End Sub

Private Sub Picture4_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\������.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\������.BMP")
Picture10.Picture = LoadPicture(App.Path & "\������.BMP")

End Sub

Private Sub Picture5_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\air.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\air��.BMP")
Picture10.Picture = LoadPicture(App.Path & "\air��.BMP")

End Sub

Private Sub Picture6_Click()
Form2.Picture1.Picture = LoadPicture(App.Path & "\Enemy.BMP")
Form2.Picture2.Picture = LoadPicture(App.Path & "\Enemy��.BMP")
Picture10.Picture = LoadPicture(App.Path & "\Enemy��.BMP")

End Sub

Private Sub Picture7_Click()

Form2.Picture9.Picture = LoadPicture(App.Path & "\���̻�����.BMP")
Form2.Picture10.Picture = LoadPicture(App.Path & "\���̻��Ϲ�.BMP")
Picture11.Picture = LoadPicture(App.Path & "\���̻��Ϲ�.BMP")

End Sub

Private Sub Picture8_Click()
Form2.Picture9.Picture = LoadPicture(App.Path & "\�̻�����.BMP")
Form2.Picture10.Picture = LoadPicture(App.Path & "\�̻��Ϲ�.BMP")
Picture11.Picture = LoadPicture(App.Path & "\�̻��Ϲ�.BMP")

End Sub

Private Sub Picture9_Click()
Form2.Picture9.Picture = LoadPicture(App.Path & "\shoot.BMP")
Form2.Picture10.Picture = LoadPicture(App.Path & "\shoot��.BMP")
Picture11.Picture = LoadPicture(App.Path & "\shoot��.BMP")

End Sub
