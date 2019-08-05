VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form mumain 
   BackColor       =   &H80000009&
   Caption         =   "타자 프로그램"
   ClientHeight    =   5715
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5565
   DrawStyle       =   3  '대시-점
   Icon            =   "타자연습 V1.0 main.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "타자연습 V1.0 main.frx":030A
   ScaleHeight     =   5715
   ScaleWidth      =   5565
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      Picture         =   "타자연습 V1.0 main.frx":0614
      ScaleHeight     =   10035
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   -120
      Width           =   14295
      Begin VB.Timer Timer12 
         Interval        =   1000
         Left            =   4080
         Top             =   4800
      End
      Begin MCI.MMControl mid11 
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   4200
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Timer Timer11 
         Interval        =   10
         Left            =   3960
         Top             =   3960
      End
      Begin VB.Timer Timer10 
         Interval        =   1000
         Left            =   120
         Top             =   4560
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   5400
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   4762
               MinWidth        =   4762
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer9 
         Interval        =   1000
         Left            =   4680
         Top             =   4440
      End
      Begin VB.Timer Timer8 
         Interval        =   1000
         Left            =   3960
         Top             =   5400
      End
      Begin VB.Timer Timer7 
         Interval        =   1000
         Left            =   1680
         Top             =   5400
      End
      Begin VB.Timer Timer6 
         Interval        =   1000
         Left            =   1080
         Top             =   5400
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   240
         Top             =   5400
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   4920
         Top             =   4440
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   4560
         Top             =   5400
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   4800
         Top             =   4800
      End
      Begin VB.Frame Frame1 
         Caption         =   "곡 선택"
         Height          =   3495
         Left            =   2640
         TabIndex        =   14
         Top             =   480
         Width           =   2295
         Begin VB.OptionButton Option11 
            Caption         =   "배경음악 9"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3120
            Width           =   1815
         End
         Begin VB.OptionButton Option10 
            Caption         =   "배경음악 8"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2760
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "배경음악 1"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "배경음악 2"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            Caption         =   "배경음악 3"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "배경음악 4"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton Option7 
            Caption         =   "배경음악 5"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   1935
         End
         Begin VB.OptionButton Option8 
            Caption         =   "배경음악 6"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2040
            Width           =   1935
         End
         Begin VB.OptionButton Option9 
            Caption         =   "배경음악 7"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2400
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "소리 켜기 / 끄기"
         Height          =   1095
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton Option5 
            Caption         =   "소리 켜기"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option6 
            Caption         =   "소리 끄기"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin MCI.MMControl mid10 
            Height          =   330
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   582
            _Version        =   393216
            DeviceType      =   ""
            FileName        =   ""
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "메뉴로"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   2
         Top             =   4200
         Width           =   5535
      End
      Begin MCI.MMControl mid3 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3240
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Timer Timer1 
         Interval        =   1250
         Left            =   1080
         Top             =   840
      End
      Begin MCI.MMControl mid9 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid8 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid7 
         Height          =   330
         Left            =   -720
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid6 
         Height          =   330
         Left            =   -600
         TabIndex        =   6
         Top             =   3720
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid5 
         Height          =   330
         Left            =   -600
         TabIndex        =   7
         Top             =   3360
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid4 
         Height          =   330
         Left            =   -840
         TabIndex        =   8
         Top             =   3000
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid2 
         Height          =   330
         Left            =   -1200
         TabIndex        =   9
         Top             =   2640
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4095
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7223
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "사운드"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mufile 
      Caption         =   "파일(&f)"
      Begin VB.Menu muso 
         Caption         =   "설정"
      End
      Begin VB.Menu muline1 
         Caption         =   "-"
      End
      Begin VB.Menu muend 
         Caption         =   "끝내기"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mugame 
      Caption         =   "타자게임(&g)"
      Begin VB.Menu mu111 
         Caption         =   "산성비"
         Shortcut        =   ^A
      End
      Begin VB.Menu musdsd 
         Caption         =   "전쟁"
         Shortcut        =   ^D
      End
      Begin VB.Menu muwkqrl 
         Caption         =   "단어잡기"
         Shortcut        =   ^G
      End
      Begin VB.Menu muekfflrl 
         Caption         =   "달리기"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu muxkwk 
      Caption         =   "타자연습(&s)"
      Begin VB.Menu muwkfl 
         Caption         =   "자리연습"
         Shortcut        =   ^Z
      End
      Begin VB.Menu murmf 
         Caption         =   "짧은글 연습"
         Shortcut        =   ^I
      End
      Begin VB.Menu mulong 
         Caption         =   "긴글 연습"
         Shortcut        =   ^L
      End
      Begin VB.Menu murjawjd 
         Caption         =   "타자 검정"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu muhelp 
      Caption         =   "도움말(&h)"
      Begin VB.Menu mume 
         Caption         =   "제작자"
      End
      Begin VB.Menu muhelp1 
         Caption         =   "도움"
      End
      Begin VB.Menu muver 
         Caption         =   "버전"
      End
   End
End
Attribute VB_Name = "mumain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim cnt As Byte
Dim CNT1 As Byte
Dim cnt2 As Byte
Dim cnt3 As Byte
Dim cnt4 As Byte
Dim cnt5 As Byte
Dim cnt6 As Byte
Dim cnt7 As Byte
Dim cnt8 As Byte
Dim aaa As Byte
Dim bbb As Byte
Dim ccc As Byte
Private Sub Command1_Click()
mid2.Visible = False
mid2.Enabled = False
mid3.Visible = False
mid3.Enabled = False
mid4.Visible = False
mid4.Enabled = False
mid5.Visible = False
mid5.Enabled = False
mid6.Visible = False
mid6.Enabled = False
mid7.Visible = False
mid7.Enabled = False
mid8.Visible = False
mid8.Enabled = False
mid9.Visible = False
mid9.Enabled = False
mid10.Visible = False
mid10.Enabled = False
mid11.Visible = False
mid11.Enabled = False

Picture1.Visible = True
Picture1.Enabled = True
Command1.Enabled = False
Command1.Visible = False
TabStrip1.Enabled = False
TabStrip1.Visible = False
Frame1.Visible = False
Frame1.Enabled = False
Frame2.Visible = False
Frame2.Enabled = False

End Sub

Private Sub Form_Load()
Timer3.Enabled = True
Timer4.Enabled = True
Timer2.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
Timer12.Enabled = True

mid2.Visible = False
mid2.Enabled = False
mid3.Visible = False
mid3.Enabled = False
mid4.Visible = False
mid4.Enabled = False
mid5.Visible = False
mid5.Enabled = False
mid6.Visible = False
mid6.Enabled = False
mid7.Visible = False
mid7.Enabled = False
mid8.Visible = False
mid8.Enabled = False
mid9.Visible = False
mid9.Enabled = False
mid10.Visible = False
mid10.Enabled = False
mid11.Visible = False
mid11.Enabled = False

Command1.Enabled = False
Command1.Visible = False
TabStrip1.Enabled = False
TabStrip1.Visible = False
Frame1.Visible = False
Frame1.Enabled = False
Frame2.Visible = False
Frame2.Enabled = False
Timer1.Enabled = False
If Option5.Value = True Then
Frame1.Enabled = True
End If
If Option6.Value = True Then
Frame1.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Option10.Enabled = False
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
End If
If Option1.Value = True Then
mid2.FileName = App.Path + "\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
End If
If Option2.Value = True Then
mid4.FileName = App.Path + "\31.mid"
mid4.Command = "open"
mid4.Command = "play"
End If
If Option3.Value = True Then
mid5.FileName = App.Path + "\11.mid"
mid5.Command = "open"
mid5.Command = "play"
End If
If Option4.Value = True Then
mid6.FileName = App.Path + "\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
End If
If Option5.Value = True Then
Frame1.Enabled = True
End If
If Option7.Value = True Then
mid7.FileName = App.Path + "\12.mid"
mid7.Command = "open"
mid7.Command = "play"
End If
If Option8.Value = True Then
mid8.FileName = App.Path + "\34.mid"
mid8.Command = "open"
mid8.Command = "play"
End If
If Option9.Value = True Then
mid9.FileName = App.Path + "\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
End If
If Option10.Value = True Then
mid10.FileName = App.Path + "\Yestrday.mid"
mid10.Command = "open"
mid10.Command = "play"
End If
mid3.FileName = App.Path + "\End.wav"
mid3.Command = "open"
If Option11.Value = True Then
mid11.FileName = App.Path + "\카탈로그.mid"
mid11.Command = "open"
mid11.Command = "play"
End If
End Sub

Private Sub mu111_Click()
Form1.Show
End Sub

Private Sub mu111help_Click()
       
       

End Sub

Private Sub muekfflrl_Click()
Form9.Show
End Sub

Private Sub muend_Click()
    
    a = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If a = vbYes Then
     mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"
     
     mid3.Command = "play"
     Timer1.Enabled = True
         Else
        Exit Sub
    End If

End Sub

Private Sub muhelp1_Click()
MsgBox "이 프로그램은 누구나 쉽게 이용할수 있게" + Chr(13) + Chr(13) _
       + "만들었습니다.^^ 그래서 도움말은 생략하겠습니다." + Chr(13) + Chr(13) _
       + "궁금하신점이 있으시면 메일로 보내주세요." + Chr(13) + Chr(13) _
       + "도움주신분 : 모레컴퓨터학원의 김선생님, 그리고 나의 가족들"
End Sub

Private Sub mulong_Click()
Form4.Show
End Sub

Private Sub mume_Click()
MsgBox "안녕하세요..^^" + Chr(13) + Chr(13) _
                  + "이 프로그램의 제작자인 최인균 이라고 합니다." + Chr(13) + Chr(13) _
                  + "비록 잘 못짠 프로그램 일지라도 잘 봐주세요." + Chr(13) + Chr(13) _
                  + "제 소개를 할께요." + Chr(13) + Chr(13) _
                  + "86년 2월22일 태어났고요, 소년입니다." + Chr(13) + Chr(13) _
                  + "그리고 현재 능곡중학교 3학년에 재학중입니다" + Chr(13) + Chr(13) _
                  + "E-Mail = heman1@hitel.net" + Chr(13) + Chr(13) _
                  + "그럼 재미있게 즐기세요..^^" + Chr(13) + Chr(13) _
                  + "1.3V 제작 : 2000년 10월"

End Sub

Private Sub murjawjd_Click()
Form5.Show
End Sub

Private Sub murmf_Click()
Form3.Show
End Sub

Private Sub musdsd_Click()
Form7.Show
End Sub

Private Sub muso_Click()
Command1.Enabled = True
Command1.Visible = True
TabStrip1.Enabled = True
TabStrip1.Visible = True
Frame1.Visible = True
Frame1.Enabled = True
Frame2.Visible = True
Frame2.Enabled = True
Picture1.Visible = False
Picture1.Visible = True
End Sub

Private Sub muver_Click()
muloag.Show
End Sub

Private Sub muwkfl_Click()
Form6.Show
End Sub

Private Sub muwkqrl_Click()
Form8.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

mid2.FileName = App.Path + "\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
End If
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid11.Command = "close"

mid10.FileName = App.Path + "\Yestrday.mid"
mid10.Command = "open"
mid10.Command = "play"
End If

End Sub

Private Sub Option11_Click()
If Option11.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"

mid11.FileName = App.Path + "\카탈로그.mid"
mid11.Command = "open"
mid11.Command = "play"
End If


End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
mid2.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

mid4.FileName = App.Path + "\31.mid"
mid4.Command = "open"
mid4.Command = "play"
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

mid5.FileName = App.Path + "\11.mid"
mid5.Command = "open"
mid5.Command = "play"
End If

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
mid4.Command = "close"
mid5.Command = "close"
mid2.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

mid6.FileName = App.Path + "\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
End If

End Sub

Private Sub Option5_Click()
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Option10.Enabled = True
Option11.Enabled = True
If Option1.Value = True Then
mid2.FileName = App.Path + "\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
End If
If Option2.Value = True Then
mid4.FileName = App.Path + "\31.mid"
mid4.Command = "open"
mid4.Command = "play"
End If
If Option3.Value = True Then
mid5.FileName = App.Path + "\11.mid"
mid5.Command = "open"
mid5.Command = "play"
End If
If Option4.Value = True Then
mid6.FileName = App.Path + "\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
End If
If Option5.Value = True Then
Frame1.Enabled = True
End If
If Option7.Value = True Then
mid7.FileName = App.Path + "\12.mid"
mid7.Command = "open"
mid7.Command = "play"
End If
If Option8.Value = True Then
mid8.FileName = App.Path + "\34.mid"
mid8.Command = "open"
mid8.Command = "play"
End If
If Option9.Value = True Then
mid9.FileName = App.Path + "\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
End If
If Option10.Value = True Then
mid10.FileName = App.Path + "\Yestrday"
mid10.Command = "open"
mid10.Command = "play"
End If
If Option11.Value = True Then
mid11.FileName = App.Path + "\카탈로그.mid"
mid11.Command = "open"
mid11.Command = "play"
End If

End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
Frame1.Enabled = False
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

End If
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Option10.Enabled = False
Option11.Enabled = False

End Sub

Private Sub Option7_Click()
If Option7.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid7.FileName = App.Path + "\12.mid"
mid7.Command = "open"
mid7.Command = "play"
End If

End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid8.FileName = App.Path + "\34.mid"
mid8.Command = "open"
mid8.Command = "play"
End If

End Sub

Private Sub Option9_Click()
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid10.Command = "close"
If Option9.Value = True Then
mid9.FileName = App.Path + "\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
End If

End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Sub Timer10_Timer()
aaa = aaa + 1
If aaa = 60 Then
bbb = bbb + 1
aaa = 0
End If
If bbb = 60 Then
ccc = ccc + 1
bbb = 0
End If
StatusBar1.Panels(1).Text = "총 연습시간: " & ccc & "시간 " & bbb & "분 " & aaa & "초 "
StatusBar1.Panels(3).Text = Format(Date, "yy-mm-dd (aaa)")
StatusBar1.Panels(2).Text = Format(Time, "hh:nn:ss")
End Sub

Private Sub Timer11_Timer()
If Option6.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid11.Command = "close"

End If

End Sub

Private Sub Timer12_Timer()
If Option11.Value = True Then
cnt8 = cnt8 + 1
If cnt8 = 65 Then
mid11.Command = "close"
mid11.Command = "open"
mid11.Command = "play"
cnt8 = 0
End If
End If

End Sub

Private Sub Timer2_Timer()
If Option1.Value = True Then
cnt = cnt + 1
If cnt = 175 Then
mid2.Command = "close"
mid2.Command = "open"
mid2.Command = "play"
cnt = 0
End If
End If
End Sub

Private Sub Timer3_Timer()
If Option2.Value = True Then
CNT1 = CNT1 + 1
If CNT1 = 120 Then
mid4.Command = "close"
mid4.Command = "open"
mid4.Command = "play"
CNT1 = 0
End If
End If

End Sub

Private Sub Timer4_Timer()
If Option3.Value = True Then
cnt2 = cnt2 + 1
If cnt2 = 155 Then
mid5.Command = "close"
mid5.Command = "open"
mid5.Command = "play"
cnt2 = 0
End If
End If
End Sub

Private Sub Timer5_Timer()
If Option4.Value = True Then
cnt3 = cnt3 + 1
If cnt3 = 100 Then
mid6.Command = "close"
mid6.Command = "open"
mid6.Command = "play"
cnt3 = 0
End If

End If
End Sub

Private Sub Timer6_Timer()
If Option7.Value = True Then
cnt4 = cnt4 + 1
If cnt4 = 60 Then
mid7.Command = "close"
mid7.Command = "open"
mid7.Command = "play"
cnt4 = 0
End If
End If
End Sub

Private Sub Timer7_Timer()
If Option8.Value = True Then
cnt5 = cnt5 + 1
If cnt5 = 120 Then
mid8.Command = "close"
mid8.Command = "open"
mid8.Command = "play"
cnt5 = 0
End If
End If
End Sub

Private Sub Timer8_Timer()
If Option9.Value = True Then
cnt6 = cnt6 + 1
If cnt6 = 250 Then
mid9.Command = "close"
mid9.Command = "open"
mid9.Command = "play"
cnt6 = 0
End If
End If
End Sub

Private Sub Timer9_Timer()
If Option10.Value = True Then
cnt7 = cnt7 + 1
If cnt7 = 120 Then
mid10.Command = "close"
mid10.Command = "open"
mid10.Command = "play"
cnt7 = 0
End If
End If
End Sub
