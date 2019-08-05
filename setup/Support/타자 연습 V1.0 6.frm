VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "자리연습"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "타자 연습 V1.0 6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   5145
   ScaleWidth      =   6585
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command4 
      Caption         =   "영어연습"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "한글연습"
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
      Caption         =   "시작하기"
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "메뉴로"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "텍스트박스에서 치시면 됩니다..그리고 만약 글자가 나오지 않으면 시작하기를 한번더 눌러주세요. "
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
      Caption         =   "맞은갯수 :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "간단한 자리연습"
      BeginProperty Font 
         Name            =   "굴림"
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
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
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
Dim 한글자리(26), a, 영어자리(27)

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Label1.Visible = True
Label4.Caption = 0
End Sub

Private Sub Command3_Click()
Label1.Visible = False
한글자리(1) = "ㄱ"
한글자리(2) = "ㄴ"
한글자리(3) = "ㄷ"
한글자리(4) = "ㄹ"
한글자리(5) = "ㅁ"
한글자리(6) = "ㅂ"
한글자리(7) = "ㅅ"
한글자리(8) = "ㅇ"
한글자리(9) = "ㅈ"
한글자리(10) = "ㅊ"
한글자리(11) = "ㅌ"
한글자리(12) = "ㅋ"
한글자리(13) = "ㅍ"
한글자리(26) = "ㅎ"
한글자리(15) = "ㅛ"
한글자리(16) = "ㅕ"
한글자리(17) = "ㅑ"
한글자리(18) = "ㅐ"
한글자리(19) = "ㅔ"
한글자리(20) = "ㅗ"
한글자리(21) = "ㅓ"
한글자리(22) = "ㅏ"
한글자리(23) = "ㅣ"
한글자리(24) = "ㅠ"
한글자리(25) = "ㅜ"
한글자리(26) = "ㅡ"
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Command3.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command4_Click()
Label1.Visible = False
영어자리(1) = "A"
영어자리(2) = "B"
영어자리(3) = "C"
영어자리(4) = "D"
영어자리(5) = "E"
영어자리(6) = "F"
영어자리(7) = "G"
영어자리(8) = "H"
영어자리(9) = "I"
영어자리(10) = "J"
영어자리(11) = "K"
영어자리(12) = "L"
영어자리(13) = "M"
영어자리(14) = "N"
영어자리(15) = "O"
영어자리(16) = "P"
영어자리(18) = "Q"
영어자리(17) = "R"
영어자리(19) = "S"
영어자리(20) = "T"
영어자리(21) = "U"
영어자리(22) = "V"
영어자리(23) = "W"
영어자리(24) = "U"
영어자리(25) = "X"
영어자리(26) = "Y"
영어자리(27) = "Z"
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Form_Load()
Randomize
Label1.Visible = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Command3.Enabled = False Then
If Label1.Caption = "ㄱ" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄴ" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄷ" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄹ" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅁ" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅂ" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅅ" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅇ" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅈ" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅊ" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅌ" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅋ" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅍ" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅎ" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "ㅛ" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "ㅕ" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅑ" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅐ" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅔ" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅗ" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅓ" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅏ" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅣ" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅠ" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅜ" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅡ" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
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
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "B" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "C" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If


If Label1.Caption = "D" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "E" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "F" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "G" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "H" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "I" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "J" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "K" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "L" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "M" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "N" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "O" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "P" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Q" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "R" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "S" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "T" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "U" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "V" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "W" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "X" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Y" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "Z" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
End If
If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If

End Sub

