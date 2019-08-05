VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "´Ü¾î Àâ±â"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "Å¸ÀÚ ¿¬½À V1.0 8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   7575
   ScaleWidth      =   8310
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "µµ¿ò¸»"
      Height          =   615
      Left            =   6840
      Style           =   1  '±×·¡ÇÈ
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
         Name            =   "±¼¸²"
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
      Caption         =   "¸Þ´º"
      Height          =   615
      Left            =   5280
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ã³À½ºÎÅÍ"
      Height          =   615
      Left            =   3600
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   18
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Áß´ÜÇÏ±â"
      Height          =   615
      Left            =   1920
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "½ÃÀÛÇÏ±â(&S)"
      Height          =   615
      Left            =   240
      Style           =   1  '±×·¡ÇÈ
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
      Caption         =   "Á¦ÇÑ ½Ã°£ :"
      BeginProperty Font 
         Name            =   "±¼¸²"
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
         Name            =   "±¼¸²"
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
      Caption         =   "¸ÂÀº °¹¼ö :"
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   6120
      Shape           =   2  'Å¸¿øÇü
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   4200
      Shape           =   2  'Å¸¿øÇü
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   2280
      Shape           =   2  'Å¸¿øÇü
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   360
      Shape           =   2  'Å¸¿øÇü
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   6120
      Shape           =   2  'Å¸¿øÇü
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   4200
      Shape           =   2  'Å¸¿øÇü
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   2280
      Shape           =   2  'Å¸¿øÇü
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   360
      Shape           =   2  'Å¸¿øÇü
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   6120
      Shape           =   2  'Å¸¿øÇü
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   4200
      Shape           =   2  'Å¸¿øÇü
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   2280
      Shape           =   2  'Å¸¿øÇü
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   360
      Shape           =   2  'Å¸¿øÇü
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   6120
      Shape           =   2  'Å¸¿øÇü
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   4200
      Shape           =   2  'Å¸¿øÇü
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   2280
      Shape           =   2  'Å¸¿øÇü
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '´Ü»ö
      Height          =   735
      Left            =   360
      Shape           =   2  'Å¸¿øÇü
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
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
Dim ³¹¸»(200), b, bb, c, d, e, f, g, h, i, k, l, m, n, o, p, q, r, cc, dd, ee, ff, gg, hh, ii, kk, ll, mm, nn, oo, pp, qq, rr, aaa, ab

Private Sub Command1_Click()
Text1.IMEMode = vbIMEModeHangul 'ÇÑ±Û
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
Label1(i).Caption = ³¹¸»(a)
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
Label1(i).Caption = ³¹¸»(a)
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
MsgBox "<´Ü¾îÀâ±â µµ¿ò¸»>" + Chr(13) + Chr(13) _
        + "¾È³çÇÏ¼¼¿ä." + Chr(13) + Chr(13) _
        + "´Ü¾îÀâ±â´Â µÎ´õÁö Àâ±â¶õ °ÔÀÓÀ» °³Á¶ÇÑ°Ì´Ï´Ù" + Chr(13) + Chr(13) _
        + "´Ü¾î°¡ ³ª¿À¸é Àß º¸½Ã°í ÃÄÁÖ½Ã±â ¹Ù¶ø´Ï´Ù. Àç¹ÌÀÖ°Ô ÇÏ¼¼¿ä.^^"
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
Label1(0).Caption = ³¹¸»(a)
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
Label1(1).Caption = ³¹¸»(a)
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
Label1(2).Caption = ³¹¸»(a)
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
Label1(3).Caption = ³¹¸»(a)
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
Label1(4).Caption = ³¹¸»(a)
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
Label1(5).Caption = ³¹¸»(a)
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
Label1(6).Caption = ³¹¸»(a)
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
Label1(7).Caption = ³¹¸»(a)
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
Label1(8).Caption = ³¹¸»(a)
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
Label1(9).Caption = ³¹¸»(a)
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
Label1(10).Caption = ³¹¸»(a)
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
Label1(11).Caption = ³¹¸»(a)
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
Label1(12).Caption = ³¹¸»(a)
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
Label1(13).Caption = ³¹¸»(a)
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
Label1(14).Caption = ³¹¸»(a)
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
Label1(15).Caption = ³¹¸»(a)
End If


End Sub

Private Sub Timer33_Timer()
aaa = aaa - 1
Label5.Caption = aaa
If Label5 = 0 Then
MsgBox "½Ã°£ÀÌ ´Ù µÇ¾ú½À´Ï´Ù."
MsgBox "´ç½ÅÀÇ ¸ÂÀº °¹¼ö´Â " & ab & "°³ ÀÔ´Ï´Ù."
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
Label1(i).Caption = ³¹¸»(a)

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
