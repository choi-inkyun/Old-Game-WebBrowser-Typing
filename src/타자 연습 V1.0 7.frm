VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form7 
   Caption         =   "ÀüÀï"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   9480
   Icon            =   "Å¸ÀÚ ¿¬½À V1.0 7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   6855
   ScaleWidth      =   9480
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
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
      Caption         =   "µµ¿ò¸»"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  '±×·¡ÇÈ
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
      Caption         =   "Áß´ÜÇÏ±â"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  '±×·¡ÇÈ
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
      Caption         =   "³­ÀÌµµ"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   1695
      Begin VB.OptionButton Option3 
         Caption         =   "ÃÊº¸ÀÚ"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Áß±ÞÀÚ"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "»ó±ÞÀÚ"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Caption         =   "¸Þ´º·Î"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   10
      Top             =   6120
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ã³À½ºÎÅÍ"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4200
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   9
      Top             =   6120
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "½ÃÀÛÇÏ±â(&S)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   360
      Style           =   1  '±×·¡ÇÈ
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
      Caption         =   "Àû±ºÀÇ º´·Â"
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
      Caption         =   "È® ÀÎ :"
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
      Caption         =   "È® ÀÎ :"
      BeginProperty Font 
         Name            =   "±¼¸²"
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
         Name            =   "±¼¸²"
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
         Name            =   "±¼¸²"
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
      Caption         =   "Á¡¼ö :"
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Caption         =   "´Ü ¾î :"
      BeginProperty Font 
         Name            =   "±¼¸²"
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
         Name            =   "±¼¸²"
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
      FillStyle       =   0  '´Ü»ö
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
Dim ³¹¸»(200), k, g, e, p, o, i, u, t, r, m, n, s, q, w, v, f, h, J, l, aa, bb, cc, dd, aaa, bbb, ccc, ddd
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
Text1.IMEMode = vbIMEModeHangul 'ÇÑ±Û
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
Label11.Caption = "µ¿°ú ¼­ÀÇ ÀüÀïÀÌ ½ÃÀÛµÇ¾ú½À´Ï´Ù. ÀûÀ» Àü¸ê½ÃÅ°´Â°Ô ¸ñÀûÀÔ´Ï´Ù. Çà¿îÀ» ºô°Ú½À´Ï´Ù."
End Sub

Private Sub Command2_Click()
Label11.Caption = "ÀüÀïÀÌ ´Ù½Ã ¹ß¹ßÇÏ·Á°í ÇÕ´Ï´Ù"
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
Label11.Caption = "ÈÞÀüÀÔ´Ï´Ù"
Label2.Visible = False
End Sub

Private Sub Command5_Click()
MsgBox "<ÀüÀï µµ¿ò¸»>" + Chr(13) + Chr(13) _
      + "¾È³çÇÏ¼¼¿ä. ÀüÀï¿¡¼­´Â Àû±¹ÀÇ º´·ÂÀ» 0 À¸·Î ¸¸µé¸é ÀÌ±â°Ô µË´Ï´Ù" + Chr(13) + Chr(13) _
      + "ÇÏÁö¸¸ Àû±ºÀÇ º´·ÂÀÌ 7800Á¤µµ°¡ µÇ¸é Áö°Ô µË´Ï´Ù" + Chr(13) + Chr(13) _
      + "¿©·¯ºÐÀº ´Ü¾î¿¡ ½áÀÖ´Â ´Ü¾î¸¦ Ä¡½Ã¸é ´Ü¾î°¡ Àû±º°ú ½Î¿ì°Ô µË´Ï´Ù" + Chr(13) + Chr(13) _
      + "ÀÌ ÀüÀï¿¡¼­´Â ÀûÀº ±º»ç·Î ½Â¸®ÇÑ°ÍÀÌ Áß¿äÇÕ´Ï´Ù" + Chr(13) + Chr(13) _
      + "Çà¿îÀ» ºô²²¿ä...Àç¹ÌÀÖ°Ô ÇÏ¼¼¿ä~.."
End Sub

Private Sub Form_Load()
³¹¸»(1) = "ÄÄÇ»ÅÍ"
³¹¸»(2) = "´ëÇÑ¹Î±¹"
³¹¸»(3) = "¹Ì³ª¸®"
³¹¸»(4) = "Å¸ÀÚ"
³¹¸»(5) = "¿öµå"
³¹¸»(6) = "°£Áö·´´Ù"
³¹¸»(7) = "¸Û°Ô"
³¹¸»(8) = "ÇØ»ï"
³¹¸»(9) = "¸»¹ÌÀß"
³¹¸»(10) = "¹è"
³¹¸»(11) = "ÀÏ¿äÀÏ"
³¹¸»(12) = "Å°º¸µå"
³¹¸»(13) = "ÇÁ¸°ÅÍ"
³¹¸»(14) = "½ºÄ³³Ê"
³¹¸»(15) = "ÇÞºû"
³¹¸»(16) = "Ãà±¸"
³¹¸»(17) = "³ó±¸"
³¹¸»(18) = "¹è±¸"
³¹¸»(19) = "ºñÄ¡"
³¹¸»(20) = "Å¹±¸"
³¹¸»(21) = "ÇÇ±¸"
³¹¸»(22) = "ÇÚµåº¼"
³¹¸»(23) = "¼³¾Ç»ê"
³¹¸»(24) = "±Ý°­»ê"
³¹¸»(25) = "¹Ì±¹"
³¹¸»(26) = "´Þ"
³¹¸»(27) = "Åä¸¶Åä"
³¹¸»(28) = "¼ö¹Ú"
³¹¸»(29) = "µþ±â"
³¹¸»(30) = "½Ã°è"
³¹¸»(31) = "ÀÏ¾î³ª´Ù"
³¹¸»(32) = "°³ÇÐ"
³¹¸»(33) = "¹æÇÐ"
³¹¸»(34) = "°ÔÀ¸¸§"
³¹¸»(35) = "´ÊÀá"
³¹¸»(36) = "¼÷Á¦"
³¹¸»(37) = "°ÔÀÓ"
³¹¸»(38) = "¸¶¿ì½º"
³¹¸»(39) = "½ºÇÇÄ¿"
³¹¸»(40) = "±×·¡ÇÈ"
³¹¸»(41) = "±Ã¼­"
³¹¸»(42) = "Ä¥ÆÇ"
³¹¸»(43) = "³ÊÅÐ¿ôÀ½"
³¹¸»(44) = "¹Ù¶÷"
³¹¸»(45) = "ºû"
³¹¸»(46) = "¾îµÒ"
³¹¸»(47) = "¹«"
³¹¸»(48) = "ºÒ"
³¹¸»(49) = "¹°"
³¹¸»(50) = "¶¥"
³¹¸»(51) = "³ë·ç"
³¹¸»(52) = "³ë¸£½º¸§"
³¹¸»(53) = "³ë±¸"
³¹¸»(54) = "»À"
³¹¸»(55) = "°ÉÀ½"
³¹¸»(56) = "°©ÀÛ½º·¹"
³¹¸»(57) = "Èñ¸Á"
³¹¸»(58) = "²Þ"
³¹¸»(59) = "Ãµ»óÃµÇÏ"
³¹¸»(60) = "È²ÅÂÀÚ"
³¹¸»(61) = "»ç°í¹æ½Ä"
³¹¸»(62) = "»ì¸²"
³¹¸»(63) = "½ÅÈ­"
³¹¸»(64) = "ÆÒ"
³¹¸»(65) = "Ã¥»ó"
³¹¸»(66) = "Ä¥ÆÇ"
³¹¸»(67) = "º¼Ææ"
³¹¸»(68) = "¸¸³âÇÊ"
³¹¸»(69) = "°øÃ¥"
³¹¸»(70) = "¾î¸Ó´Ï"
³¹¸»(71) = "¾Æ¹öÁö"
³¹¸»(72) = "»ïÃÌ"
³¹¸»(73) = "°í¸ð"
³¹¸»(74) = "ÀÌ¸ð"
³¹¸»(75) = "ÇÒ¸Ó´Ï"
³¹¸»(76) = "ÇÒ¾Æ¹öÁö"
³¹¸»(77) = "µµ½Ã"
³¹¸»(78) = "¹ö½º"
³¹¸»(79) = "ÁöÇÏÃ¶"
³¹¸»(80) = "½ÄÃÊ"
³¹¸»(81) = "¼³ÅÁ"
³¹¸»(82) = "¼Ò±Ý"
³¹¸»(83) = "³ªÆ®·ý"
³¹¸»(84) = "±Ý"
³¹¸»(85) = "½Ä¹°"
³¹¸»(86) = "µ¿¹°"
³¹¸»(87) = "°í¾çÀÌ"
³¹¸»(88) = "»çÀÚ"
³¹¸»(89) = "È£¶ûÀÌ"
³¹¸»(90) = "ÇÏÀÌ¿¡³ª"
³¹¸»(91) = "ÀÎÅÍ³Ý"
³¹¸»(92) = "³×Æ®¿öÅ©"
³¹¸»(93) = "¼­·ù"
³¹¸»(94) = "ÈÞÁöÅë"
³¹¸»(95) = "±èÄ¡"
³¹¸»(96) = "»§"
³¹¸»(97) = "¸¸È­"
³¹¸»(98) = "½Ç·ÎÆù"
³¹¸»(99) = "¹ÙÀÌ¿Ã¸°"
³¹¸»(100) = "¶±"
³¹¸»(101) = "¾È°æ"
³¹¸»(102) = "½Ò"
³¹¸»(103) = "±³°ú¼­"
³¹¸»(104) = "¼Ò¼³"
³¹¸»(105) = "Ã¥"
³¹¸»(106) = "¿µÈ­"
³¹¸»(107) = "º¸¸®"
³¹¸»(108) = "ÀÇÀÚ"
³¹¸»(109) = "°É»ó"
³¹¸»(110) = "µð½ºÄÏ"
³¹¸»(111) = "Âü°í¼­"
³¹¸»(112) = "½ÃÇè"
³¹¸»(113) = "ÀÚ°ÝÁõ"
³¹¸»(114) = "¾ÆÀÌ°í"
³¹¸»(115) = "¿¬½À"
³¹¸»(116) = "´ëÈ¸"
³¹¸»(117) = "ÅÚ·¹ºñÀü"
³¹¸»(118) = "¹æÁ¤¸Â´Ù"
³¹¸»(119) = "¾Æ½Ã¾Æ"
³¹¸»(120) = "¼¼°è"
³¹¸»(121) = "¾Ö±¹°¡"
³¹¸»(122) = "ÇÏ´À´Ô"
³¹¸»(123) = "¹«±ÃÈ­"
³¹¸»(124) = "»ê"
³¹¸»(125) = "¾îÁö·´´Ù"
³¹¸»(126) = "È²´çÇÏ´Ù"
³¹¸»(127) = "ÃÑ"
³¹¸»(128) = "»ìÀÎ"
³¹¸»(129) = "Á×´Ù"
³¹¸»(130) = "°­µµ"
³¹¸»(131) = "¹üÁË"
³¹¸»(132) = "¾ö¸¶"
³¹¸»(133) = "¾Æºü"
³¹¸»(134) = "¾ó´Ù"
³¹¸»(135) = "µ¿»ó"
³¹¸»(136) = "°íµå¸§"
³¹¸»(137) = "´«"
³¹¸»(138) = "ºñ"
³¹¸»(139) = "¿¡¸Þ¶öµå"
³¹¸»(140) = "´ÙÀÌ¾Æ"
³¹¸»(141) = "Ä«µå"
³¹¸»(142) = "º½"
³¹¸»(143) = "¿©¸§"
³¹¸»(144) = "°¡À»"
³¹¸»(145) = "°Ü¿ï"
³¹¸»(146) = "ÃáÇÏÃßµ¿"
³¹¸»(147) = "³í"
³¹¸»(148) = "¹ç"
³¹¸»(149) = "Ãß¼ö"
³¹¸»(150) = "¾ÆÄ§"
³¹¸»(151) = "Á¡½É"
³¹¸»(152) = "Àú³á"
³¹¸»(153) = "¹ã"
³¹¸»(154) = "¿­¼è"
³¹¸»(155) = "±¸µÎ"
³¹¸»(156) = "½Å¹ß"
³¹¸»(157) = "¿îµ¿È­"
³¹¸»(158) = "º£ÀÌÁ÷"
³¹¸»(159) = "¸ðÀÓ"
³¹¸»(160) = "µ¿¾Æ¸®"
³¹¸»(161) = "µ¿È£È¸"
³¹¸»(162) = "¿µÀå"
³¹¸»(163) = "Ä®"
³¹¸»(164) = "¿µ¾î"
³¹¸»(165) = "¸í¾ð"
³¹¸»(166) = "ÇÏ´Ã"
³¹¸»(167) = "»ç¶û"
³¹¸»(168) = "±â¾ïÇÏ´Ù"
³¹¸»(169) = "¿µÈ¥"
³¹¸»(170) = "°úÀÚ"
³¹¸»(171) = "»çÀü"
³¹¸»(172) = "ºÒ°¡´É"
³¹¸»(173) = "¹Ð·¹´Ï¾ö"
³¹¸»(174) = "¿µÀç"
³¹¸»(175) = "¾ê±â"
³¹¸»(176) = "¿À·¡µÇ´Ù"
³¹¸»(177) = "¿Àµµ¹æÁ¤"
³¹¸»(178) = "¿ô´Ù"
³¹¸»(179) = "¹°"
³¹¸»(180) = "¿Õ±Ã"
³¹¸»(181) = "±ÍÁ·"
³¹¸»(182) = "¾ç¹Ý"
³¹¸»(183) = "µ·"
³¹¸»(184) = "¿Ü·Ó´Ù"
³¹¸»(185) = "¿Ü°¡´ì"
³¹¸»(186) = "¹ú·¹"
³¹¸»(187) = "¿ì¹°"
³¹¸»(188) = "¿ìµî»ý"
³¹¸»(189) = "¿ìµÎ¸Ó¸®"
³¹¸»(190) = "¿ìÀ¯"
³¹¸»(191) = "¿îµ¿"
³¹¸»(192) = "¿øµÎ¸·"
³¹¸»(193) = "¿ø±Ù°¨"
³¹¸»(194) = "À§±â"
³¹¸»(195) = "¿ù½Ä"
³¹¸»(196) = "»ý¸í"
³¹¸»(197) = "³ª¹«"
³¹¸»(198) = "À¯¾ð"
³¹¸»(199) = "ÅõÀÚ"
³¹¸»(200) = "ÇÑ°á°°´Ù"


a = Int(Rnd(1) * 200)
Label2.Caption = ³¹¸»(a)
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
    Label2.Caption = ³¹¸»(a)
    If Label2.Caption = "" Then
    a = Int(Rnd(1) * 200)
    Label2.Caption = ³¹¸»(a)
     End If
    g = k * 80
    Label6.Caption = "¸Â¾Ò´Ù"
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
    Label2.Caption = ³¹¸»(a)
    If Label2.Caption = "" Then
     a = Int(Rnd(1) * 200)
     Label2.Caption = ³¹¸»(a)
    End If
    Label6.Caption = "Æ²·È´Ù"
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
MsgBox "ÀüÀïÀÌ ³¡³µ½À´Ï´Ù"
MsgBox "´ç½ÅÀº Àß ½Î¿ö ÁÖ¾úÁö¸¸ ÀüÀï¿¡¼­´Â ÆÐÇÏ°í ¸»¾Ò½À´Ï´Ù." + Chr(13) + Chr(13) _
       + "´ç½ÅÀº Àû±º¿¡°Ô " & e & " ¸¸Å­ÀÇ ÇÇÇØ¸¦ ÀÔÇû½À´Ï´Ù"
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
Label11.Caption = "´ç½ÅÀº ÀüÀï¿¡¼­ ÆÐ¹èÇÏ¿´½À´Ï´Ù"
End If

If Shape1.Left <= 0 Then
MsgBox "ÃàÇÏÇÕ´Ï´Ù.^^"
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
MsgBox "ÀüÀïÀÌ ³¡³µ½À´Ï´Ù"
MsgBox "ÃàÇÏÇÕ´Ï´Ù. ´ç½ÅÀº ÀüÀï¿¡¼­ ½Â¸®ÇÏ¿´½À´Ï´Ù" + Chr(13) + Chr(13) _
        + "´ç½ÅÀº " & e & "¸íÀÇ ±º»ç·Î ½Â¸®ÇÏ¼Ì½À´Ï´Ù"
Label5.Caption = 0
Label2.Visible = False
Frame1.Visible = True
Frame1.Enabled = True
Label11.Caption = "ÃàÇÏÇÕ´Ï´Ù. ´ç½ÅÀº ÀüÀï¿¡¼­ ½Â¸®ÇÏ¿´½À´Ï´Ù"
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
               Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

       Label11.Caption = "Àû±º¿¡¼­ " & tttt & "¸íÀÇ ¿ø±ºÀÌ µµÂøÇß½À´Ï´Ù"
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

       Label11.Caption = "Àû±º¿¡¼­ " & tttt & "¸íÀÇ ¿ø±ºÀÌ µµÂøÇß½À´Ï´Ù"
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

       Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
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

        Label11.Caption = "Àû±ºÀÇ º´»ç¸¦ " & g & " ¸í ¸¸Å­ ¼Õ½Ç½ÃÄ×½À´Ï´Ù"
       Label1(s).Left = 7800
        Label1(s).Caption = ""
        p = False
    End If
End If

End Sub
