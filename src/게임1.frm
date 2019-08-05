VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "air"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "게임1.frx":0000
   LinkMode        =   1  '원본
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   9570
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   53
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":1F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":2E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":3C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":4A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":57A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":6606
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":72FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":801E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":8CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":98F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   53
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":A616
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":B4B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":C25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":D10E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":DF8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":ED6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":FAAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":10912
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":11606
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":1232A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":12FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "게임1.frx":13BF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MCI.MMControl midi1 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   4680
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   7320
      Width           =   10020
      Begin VB.Image Image1 
         Height          =   480
         Left            =   9000
         Picture         =   "게임1.frx":14916
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "1 스테이지"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub play()
If lefts = 1 Then
xxx = xxx - 15
 
If xxx < 0 Then
xxx = 0
End If
End If
If rights = 1 Then
xxx = xxx + 15
 

End If
If ups = 1 Then
yyy = yyy - 15
 

If yyy < 0 Then
yyy = 0
End If
End If
If downs = 1 Then
yyy = yyy + 15
 

End If
    If ssss = 1 Then
    altkdlfy1 = altkdlfy1 - 50
          SUCCESS = BitBlt(Form1.hDC, altkdlfx1, altkdlfy1, Form2.Picture10.ScaleWidth, Form2.Picture10.ScaleHeight, Form2.Picture10.hDC, 0, 0, SRCAND)
          SUCCESS = BitBlt(Form1.hDC, altkdlfx1, altkdlfy1, Form2.Picture9.ScaleWidth, Form2.Picture9.ScaleHeight, Form2.Picture9.hDC, 0, 0, SRCPAINT)
End If
 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture2.ScaleWidth, Form2.Picture2.ScaleHeight, Form2.Picture2.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture1.ScaleWidth, Form2.Picture1.ScaleHeight, Form2.Picture1.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     If KeyCode = vbKeySpace Then sss = 1
     If KeyCode = 37 Then lefts = 1
     If KeyCode = 38 Then ups = 1
     If KeyCode = 39 Then rights = 1
     If KeyCode = 40 Then downs = 1
     If KeyCode = 17 And Image1.Visible = True Then
vhrvkfx1 = upwjrx1
vhrvkfy1 = upwjry1
vhrvkfx2 = upwjrx3
vhrvkfy2 = upwjry3
vhrvkfx3 = leftwjrx1
vhrvkfy3 = leftwjry1
vhrvkfx4 = rightwjrx1
vhrvkfy4 = rightwjry1
Call timer66
Call timer77
Call timer88
Call timer99
Call timer100

upwjry1 = -10
upwjrx1 = 0
upwjrcnt1 = 0
upwjry2 = -10
upwjrx2 = 0
upwjrcnt2 = 0
upwjry3 = -10
upwjrx3 = 0
upwjrcnt3 = 0
upwjry4 = -10
upwjrx4 = 0
upwjrcnt4 = 0
upwjry5 = -10
upwjrx5 = 0
upwjrcnt5 = 0

upwjrspeed1 = 5
upwjrspeed2 = 5
upwjrspeed3 = 5
upwjrspeed4 = 5
upwjrspeed5 = 5

leftwjry1 = 200
leftwjrx1 = -50
leftwjrcnt1 = 0
leftwjrspeed1 = 5
leftwjry2 = 200
leftwjrx2 = -50
leftwjrcnt2 = 0
leftwjrspeed2 = 5

rightwjry1 = 200
rightwjrx1 = 600
rightwjrcnt1 = 0
rightwjrspeed1 = 5

rightwjry2 = 300
rightwjrx2 = 700
rightwjrcnt2 = 0
rightwjrspeed2 = 5
upwjraltkdlfx1 = 1000
upwjraltkdlfy1 = 0
upwjraltkdlfx2 = 1000
upwjraltkdlfy2 = 0
upwjraltkdlfx3 = 1000
upwjraltkdlfy3 = 0
upwjraltkdlfx4 = 1000
upwjraltkdlfy4 = 0
upwjraltkdlfx5 = 1000
upwjraltkdlfy5 = 0
Image1.Visible = False
     End If
     End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If Timer1.Enabled = 0 Then Timer1.Enabled = 1
     If KeyAscii = 27 Then End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = 17 Then FIRESW = 0
     If KeyCode = 37 Then lefts = 0
     If KeyCode = 38 Then ups = 0
     If KeyCode = 39 Then rights = 0
     If KeyCode = 40 Then downs = 0

End Sub

Private Sub Form_Load()
Label4.Visible = False
midi1.Visible = False

     Randomize Timer
Call starts
     Form3.Picture1.Picture = LoadPicture(App.Path & "\back3.gif")
     
     
     Picture = LoadPicture(App.Path & "\back3.gif")

 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture2.ScaleWidth, Form2.Picture2.ScaleHeight, Form2.Picture2.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture1.ScaleWidth, Form2.Picture1.ScaleHeight, Form2.Picture1.hDC, 0, 0, SRCPAINT)
Call timer11

End Sub

Private Sub Timer1_Timer()
If Label3.Caption = "1 스테이지" Then Call timer11
If Label3.Caption = "2 스테이지" Then Call timer22
If Label3.Caption = "왕이다" Then Call timer33
If stertime = 800 Then Call timer22
'SUCCESS = BitBlt(Form1.hDC, 0, 0, Form2.Picture19.ScaleWidth, Form2.Picture19.ScaleHeight, Form2.Picture19.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub starts()
sss = 0
lefts = 0
rights = 0
ups = 0
downs = 0
BACKGROUNDCTR = 450
xxx = 270
yyy = 400
wjatn = 0
ahrt = 3
starttime = 0
upwjry1 = -10
upwjrx1 = 0
upwjrcnt1 = 0
upwjry2 = -10
upwjrx2 = 0
upwjrcnt2 = 0
upwjry3 = -10
upwjrx3 = 0
upwjrcnt3 = 0
upwjry4 = -10
upwjrx4 = 0
upwjrcnt4 = 0
upwjry5 = -10
upwjrx5 = 0
upwjrcnt5 = 0

upwjrspeed1 = 5
upwjrspeed2 = 5
upwjrspeed3 = 5
upwjrspeed4 = 5
upwjrspeed5 = 5

leftwjry1 = 200
leftwjrx1 = -50
leftwjrcnt1 = 0
leftwjrspeed1 = 5
leftwjry2 = 200
leftwjrx2 = -50
leftwjrcnt2 = 0
leftwjrspeed2 = 5

rightwjry1 = 200
rightwjrx1 = 600
rightwjrcnt1 = 0
rightwjrspeed1 = 5

rightwjry2 = 300
rightwjrx2 = 700
rightwjrcnt2 = 0
rightwjrspeed2 = 5

ssss = 0
altkdlfx1 = -50
altkdlfy1 = -50

upwjraltkdlfx1 = 1000
upwjraltkdlfy1 = 0
upwjraltkdlfx2 = 1000
upwjraltkdlfy2 = 0
upwjraltkdlfx3 = 1000
upwjraltkdlfy3 = 0
upwjraltkdlfx4 = 1000
upwjraltkdlfy4 = 0
upwjraltkdlfx5 = 1000
upwjraltkdlfy5 = 0

midi1.FileName = App.Path + "\qorud.mid"
midi1.Command = "open"
midi1.Command = "play"
wjrx = 1000
wjrt = 1000
wjrcnt = 0
vlftkfrlx = 200
vlftkfrly = -500
vlftkfrl1 = 0

End Sub

Private Sub upwjr1()

 
upwjrcnt1 = upwjrcnt1 + 1
If upwjrcnt1 = 1 Then
    Randomize Timer
upwjrx1 = Int(Rnd(10) * 550)
upwjry1 = Int(Rnd(-30) * -200)
upwjrspeed1 = Int(Rnd(10) * 25)
End If
upwjry1 = upwjry1 + upwjrspeed1
SUCCESS = BitBlt(Form1.hDC, upwjrx1, upwjry1, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, upwjrx1, upwjry1, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)
If upwjry1 > 570 Then
upwjry1 = Int(Rnd(-30) * -200)
upwjrcnt1 = 0
End If

upwjrcnt2 = upwjrcnt2 + 1
If upwjrcnt2 = 1 Then
Randomize Timer
upwjrx2 = Int(Rnd(4) * 530)
upwjry2 = Int(Rnd(-10) * -400)
upwjrspeed2 = Int(Rnd(11) * 13)
End If
upwjry2 = upwjry2 + upwjrspeed2
 SUCCESS = BitBlt(Form1.hDC, upwjrx2, upwjry2, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, upwjrx2, upwjry2, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)
If upwjry2 > 570 Then
upwjry2 = Int(Rnd(-30) * -200)
upwjrcnt2 = 0
End If

upwjrcnt5 = upwjrcnt5 + 1
If upwjrcnt5 = 1 Then
Randomize Timer
upwjrx5 = Int(Rnd(20) * 510)
upwjry5 = Int(Rnd(-30) * -300)
upwjrspeed5 = Int(Rnd(8) * 15)
End If
upwjry5 = upwjry5 + upwjrspeed5
 SUCCESS = BitBlt(Form1.hDC, upwjrx5, upwjry5, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, upwjrx5, upwjry5, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)
If upwjry5 > 570 Then
upwjry5 = Int(Rnd(-30) * -200)
upwjrcnt5 = 0
End If

End Sub
Private Sub upwjr2()

upwjrcnt3 = upwjrcnt3 + 1
If upwjrcnt3 = 1 Then
Randomize Timer
upwjrx3 = Int(Rnd(5) * 555)
upwjry3 = Int(Rnd(-50) * -300)
upwjrspeed3 = Int(Rnd(9) * 14)
End If
upwjry3 = upwjry3 + upwjrspeed3
If yyy > upwjry3 Then
    If xxx > upwjrx3 Then
    upwjrx3 = upwjrx3 + upwjrspeed3 / 2
    
    End If
    If xxx < upwjrx3 Then
    upwjrx3 = upwjrx3 - upwjrspeed3 / 2
   
    End If
End If
 SUCCESS = BitBlt(Form1.hDC, upwjrx3, upwjry3, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, upwjrx3, upwjry3, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)

If upwjry3 > 570 Then
upwjry3 = Int(Rnd(-30) * -200)
upwjrcnt3 = 0
End If

upwjrcnt4 = upwjrcnt4 + 1
If upwjrcnt4 = 1 Then
Randomize Timer
upwjrx4 = Int(Rnd(1) * 559)
upwjry4 = Int(Rnd(-20) * -150)
upwjrspeed4 = Int(Rnd(10) * 15)
End If
upwjry4 = upwjry4 + upwjrspeed4
If yyy > upwjry4 Then
    If xxx > upwjrx4 Then
    upwjrx4 = upwjrx4 + upwjrspeed4 / 2
    
    End If
    If xxx < upwjrx4 Then
    upwjrx4 = upwjrx4 - upwjrspeed4 / 2
   
    End If
End If
 
 SUCCESS = BitBlt(Form1.hDC, upwjrx4, upwjry4, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, upwjrx4, upwjry4, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)
If upwjry4 > 570 Then
upwjry4 = Int(Rnd(-30) * -200)
upwjrcnt4 = 0
End If


End Sub
Private Sub leftwjr()
leftwjrcnt1 = leftwjrcnt1 + 1
If leftwjrcnt1 = 1 Then
Randomize Timer
leftwjrx1 = Int(Rnd(-10) * -100)
leftwjry1 = Int(Rnd(50) * 200)
leftwjrspeed1 = Int(Rnd(12) * 14)
End If
leftwjrx1 = leftwjrx1 + leftwjrspeed1

SUCCESS = BitBlt(Form1.hDC, leftwjrx1, leftwjry1, Form2.Picture6.ScaleWidth, Form2.Picture6.ScaleHeight, Form2.Picture6.hDC, 0, 0, SRCAND)
SUCCESS = BitBlt(Form1.hDC, leftwjrx1, leftwjry1, Form2.Picture5.ScaleWidth, Form2.Picture5.ScaleHeight, Form2.Picture5.hDC, 0, 0, SRCPAINT)
If leftwjrx1 > 600 Then
leftwjry1 = Int(Rnd(50) * 200)
leftwjrx1 = Int(Rnd(-10) * -100)
leftwjrcnt1 = 0
End If

leftwjrcnt2 = leftwjrcnt2 + 1
If leftwjrcnt2 = 1 Then
Randomize Timer
leftwjrx2 = Int(Rnd(-10) * -100)
leftwjry2 = Int(Rnd(1) * 300)
leftwjrspeed2 = Int(Rnd(8) * 14)



End If
leftwjrx2 = leftwjrx2 + leftwjrspeed2
If xxx > leftwjrx1 Then
    If yyy > leftwjry1 Then
    leftwjry1 = leftwjry1 + leftwjrspeed1 / 2
    
    End If
    If yyy < leftwjry1 Then
    leftwjry1 = leftwjry1 - leftwjrspeed1 / 2
   
    End If
End If
 
 SUCCESS = BitBlt(Form1.hDC, leftwjrx2, leftwjry2, Form2.Picture6.ScaleWidth, Form2.Picture6.ScaleHeight, Form2.Picture6.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, leftwjrx2, leftwjry2, Form2.Picture5.ScaleWidth, Form2.Picture5.ScaleHeight, Form2.Picture5.hDC, 0, 0, SRCPAINT)
If leftwjrx2 > 602 Then
leftwjry2 = Int(Rnd(50) * 300)
leftwjrx2 = Int(Rnd(-20) * -110)
leftwjrcnt2 = 0
End If


End Sub
Private Sub rightwjr()
rightwjrcnt1 = rightwjrcnt1 + 1
If rightwjrcnt1 = 1 Then
Randomize Timer
rightwjrx1 = 650 'Int(Rnd(600) * 700)
rightwjry1 = Int(Rnd(10) * 420)
rightwjrspeed1 = Int(Rnd(10) * 13)
End If
rightwjrx1 = rightwjrx1 - rightwjrspeed1

SUCCESS = BitBlt(Form1.hDC, rightwjrx1, rightwjry1, Form2.Picture8.ScaleWidth, Form2.Picture8.ScaleHeight, Form2.Picture8.hDC, 0, 0, SRCAND)
SUCCESS = BitBlt(Form1.hDC, rightwjrx1, rightwjry1, Form2.Picture7.ScaleWidth, Form2.Picture7.ScaleHeight, Form2.Picture7.hDC, 0, 0, SRCPAINT)
If rightwjrx1 < -100 Then
rightwjry1 = Int(Rnd(10) * 420)
rightwjrx1 = 650 'Int(Rnd(600) * 700)
rightwjrcnt1 = 0
End If

rightwjrcnt2 = rightwjrcnt2 + 1
If rightwjrcnt2 = 1 Then
Randomize Timer
rightwjrx2 = 650 'Int(Rnd(750) * 800)
rightwjry2 = Int(Rnd(30) * 350)
rightwjrspeed2 = Int(Rnd(7) * 15)
End If
rightwjrx2 = rightwjrx2 - rightwjrspeed2
If xxx < rightwjrx2 Then
    If yyy > rightwjry2 Then
    rightwjry2 = rightwjry2 + rightwjrspeed2 / 2
    
    End If
    If yyy < rightwjry2 Then
    rightwjry2 = rightwjry2 - rightwjrspeed2 / 2
   
    End If
End If


SUCCESS = BitBlt(Form1.hDC, rightwjrx2, rightwjry2, Form2.Picture8.ScaleWidth, Form2.Picture8.ScaleHeight, Form2.Picture8.hDC, 0, 0, SRCAND)
SUCCESS = BitBlt(Form1.hDC, rightwjrx2, rightwjry2, Form2.Picture7.ScaleWidth, Form2.Picture7.ScaleHeight, Form2.Picture7.hDC, 0, 0, SRCPAINT)
If rightwjrx2 < -100 Then
rightwjry2 = Int(Rnd(30) * 350)
rightwjrx2 = 650 'Int(Rnd(750) * 800)
rightwjrcnt2 = 0
End If

End Sub



Private Sub Timer10_Timer()


End Sub

Private Sub Timer4_Timer()

End Sub



Private Sub cndehf()
     If xxx > upwjrx1 - 30 Then
          If xxx < upwjrx1 + 30 Then
               If yyy > upwjry1 - 30 Then
                    If yyy < upwjry1 + 30 Then

vhrvkfx1 = upwjrx1
vhrvkfy1 = upwjry1
Call timer66
upwjry1 = 568
ahrt = ahrt - 1
wjatn = wjatn + 1
                    End If
                End If
         End If
    End If

     If xxx > upwjrx2 - 30 Then
          If xxx < upwjrx2 + 30 Then
               If yyy > upwjry2 - 30 Then
                    If yyy < upwjry2 + 30 Then

vhrvkfx1 = upwjrx2
vhrvkfy1 = upwjry2
Call timer66

upwjry2 = 568
ahrt = ahrt - 1
wjatn = wjatn + 1
                    End If
                End If
         End If
    End If

     If xxx > upwjrx3 - 30 Then
          If xxx < upwjrx3 + 30 Then
               If yyy > upwjry3 - 30 Then
                    If yyy < upwjry3 + 30 Then

vhrvkfx1 = upwjrx3
vhrvkfy1 = upwjry3
Call timer66

upwjry3 = 568
ahrt = ahrt - 1
wjatn = wjatn + 1
                    End If
                End If
         End If
    End If

     If xxx > upwjrx4 - 30 Then
          If xxx < upwjrx4 + 30 Then
               If yyy > upwjry4 - 30 Then
                    If yyy < upwjry4 + 30 Then
vhrvkfx1 = upwjrx4
vhrvkfy1 = upwjry4
Call timer66

upwjry4 = 568
ahrt = ahrt - 1
wjatn = wjatn + 1
                    End If
                End If
         End If
    End If

     If xxx > upwjrx5 - 30 Then
          If xxx < upwjrx5 + 30 Then
               If yyy > upwjry5 - 30 Then
                    If yyy < upwjry5 + 30 Then

vhrvkfx1 = upwjrx5
vhrvkfy1 = upwjry5
Call timer66

upwjry5 = 568
ahrt = ahrt - 1
wjatn = wjatn + 1
                    End If
                End If
         End If
    End If

     If xxx > leftwjrx1 - 30 Then
          If xxx < leftwjrx1 + 30 Then
               If yyy > leftwjry1 - 30 Then
                    If yyy < leftwjry1 + 30 Then

vhrvkfx1 = leftwjrx1
vhrvkfy1 = leftwjry1
Call timer66

leftwjrx1 = 630
ahrt = ahrt - 1
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If xxx > leftwjrx2 - 30 Then
          If xxx < leftwjrx2 + 30 Then
               If yyy > leftwjry2 - 30 Then
                    If yyy < leftwjry2 + 30 Then

vhrvkfx1 = leftwjrx2
vhrvkfy1 = leftwjry2
Call timer66

leftwjrx2 = 630
ahrt = ahrt - 1
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If xxx > rightwjrx1 - 30 Then
          If xxx < rightwjrx1 + 30 Then
               If yyy > rightwjry1 - 30 Then
                    If yyy < rightwjry1 + 30 Then

vhrvkfx1 = rightwjrx1
vhrvkfy1 = rightwjry1
Call timer66

rightwjrx1 = -50
ahrt = ahrt - 1
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If xxx > rightwjrx2 - 30 Then
          If xxx < rightwjrx2 + 30 Then
               If yyy > rightwjry2 - 30 Then
                    If yyy < rightwjry2 + 30 Then

vhrvkfx1 = rightwjrx2
vhrvkfy1 = rightwjry2
Call timer66

rightwjrx2 = -50
ahrt = ahrt - 1
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If upwjrx1 > altkdlfx1 - 30 Then
          If upwjrx1 < altkdlfx1 + 30 Then
               If upwjry1 > altkdlfy1 - 30 Then
                    If upwjry1 < altkdlfy1 + 30 Then

vhrvkfx2 = upwjrx1
vhrvkfy2 = upwjry1
Call timer77

altkdlfx1 = -30
upwjry1 = 568
wjatn = wjatn + 1
                    
                    End If
                End If
         End If
    End If

     If upwjrx2 > altkdlfx1 - 30 Then
          If upwjrx2 < altkdlfx1 + 30 Then
               If upwjry2 > altkdlfy1 - 30 Then
                    If upwjry2 < altkdlfy1 + 30 Then

vhrvkfx2 = upwjrx2
vhrvkfy2 = upwjry2
Call timer77
altkdlfx1 = -30
upwjry2 = 568
wjatn = wjatn + 1
                    
                    End If
                End If
         End If
    End If

     If upwjrx3 > altkdlfx1 - 30 Then
          If upwjrx3 < altkdlfx1 + 30 Then
               If upwjry3 > altkdlfy1 - 30 Then
                    If upwjry3 < altkdlfy1 + 30 Then

vhrvkfx2 = upwjrx3
vhrvkfy2 = upwjry3
Call timer77
altkdlfx1 = -30
upwjry3 = 568
wjatn = wjatn + 1
                    
                    End If
                End If
         End If
    End If

     If upwjrx4 > altkdlfx1 - 30 Then
          If upwjrx4 < altkdlfx1 + 30 Then
               If upwjry4 > altkdlfy1 - 30 Then
                    If upwjry4 < altkdlfy1 + 30 Then

vhrvkfx2 = upwjrx4
vhrvkfy2 = upwjry4
Call timer77
altkdlfx1 = -30
upwjry4 = 568
wjatn = wjatn + 1
                    
                    End If
                End If
         End If
    End If

     If upwjrx5 > altkdlfx1 - 30 Then
          If upwjrx5 < altkdlfx1 + 30 Then
               If upwjry5 > altkdlfy1 - 30 Then
                    If upwjry5 < altkdlfy1 + 30 Then
vhrvkfx2 = upwjrx5
vhrvkfy2 = upwjry5
Call timer77
altkdlfx1 = -30
upwjry5 = 568
wjatn = wjatn + 1
                    
                    End If
                End If
         End If
    End If

     If leftwjrx1 > altkdlfx1 - 30 Then
          If leftwjrx1 < altkdlfx1 + 30 Then
               If leftwjry1 > altkdlfy1 - 30 Then
                    If leftwjry1 < altkdlfy1 + 30 Then

vhrvkfx5 = leftwjrx1
vhrvkfy5 = leftwjry1
Call timer100
altkdlfx1 = -50
leftwjrx1 = 630
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If leftwjrx2 > altkdlfx1 - 30 Then
          If leftwjrx2 < altkdlfx1 + 30 Then
               If leftwjry2 > altkdlfy1 - 30 Then
                    If leftwjry2 < altkdlfy1 + 30 Then

vhrvkfx5 = leftwjrx2
vhrvkfy5 = leftwjry2
Call timer100

altkdlfx1 = -50
leftwjrx2 = 630
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If






     If rightwjrx1 > altkdlfx1 - 30 Then
          If rightwjrx1 < altkdlfx1 + 30 Then
               If rightwjry1 > altkdlfy1 - 30 Then
                    If rightwjry1 < altkdlfy1 + 30 Then
vhrvkfx5 = rightwjrx1
vhrvkfy5 = rightwjry1
Call timer100

altkdlfx1 = -50
rightwjrx1 = -30
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If rightwjrx2 > altkdlfx1 - 30 Then
          If rightwjrx2 < altkdlfx1 + 30 Then
               If rightwjry2 > altkdlfy1 - 30 Then
                    If rightwjry2 < altkdlfy1 + 30 Then
vhrvkfx5 = rightwjrx2
vhrvkfy5 = rightwjry2
Call timer100

altkdlfx1 = -50
rightwjrx2 = -30
wjatn = wjatn + 2
                    End If
                End If
         End If
    End If

     If xxx > upwjraltkdlfx1 - 30 Then
          If xxx < upwjraltkdlfx1 + 30 Then
               If yyy > upwjraltkdlfy1 - 30 Then
                    If yyy < upwjraltkdlfy1 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99

ahrt = ahrt - 1
upwjraltkdlfx1 = upwjrx1
upwjraltkdlfy1 = upwjry1
                    End If
                End If
         End If
    End If

     If xxx > upwjraltkdlfx2 - 30 Then
          If xxx < upwjraltkdlfx2 + 30 Then
               If yyy > upwjraltkdlfy2 - 30 Then
                    If yyy < upwjraltkdlfy2 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
upwjraltkdlfx2 = upwjrx2
upwjraltkdlfy2 = upwjry2
                    End If
                End If
         End If
    End If


     If xxx > upwjraltkdlfx3 - 30 Then
          If xxx < upwjraltkdlfx3 + 30 Then
               If yyy > upwjraltkdlfy3 - 30 Then
                    If yyy < upwjraltkdlfy3 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
upwjraltkdlfx3 = upwjrx3
upwjraltkdlfy3 = upwjry3
                    End If
                End If
         End If
    End If

     If xxx > upwjraltkdlfx4 - 30 Then
          If xxx < upwjraltkdlfx4 + 30 Then
               If yyy > upwjraltkdlfy4 - 30 Then
                    If yyy < upwjraltkdlfy4 + 30 Then
vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
upwjraltkdlfx4 = upwjrx4
upwjraltkdlfy4 = upwjry4
                    End If
                End If
         End If
    End If

     If xxx > upwjraltkdlfx5 - 30 Then
          If xxx < upwjraltkdlfx5 + 30 Then
               If yyy > upwjraltkdlfy5 - 30 Then
                    If yyy < upwjraltkdlfy5 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
upwjraltkdlfx5 = upwjrx5
upwjraltkdlfy5 = upwjry5
                    End If
                End If
         End If
    End If

     If xxx > wjraltkdlfx - 30 Then
          If xxx < wjraltkdlfx + 30 Then
               If yyy > wjraltkdlfy - 30 Then
                    If yyy < wjraltkdlfy + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
wjraltkdlfx = wjrx
wjraltkdlfy = wjry
                    End If
                End If
         End If
    End If

    If xxx > wjrx - 30 Then
          If xxx < wjrx + 30 Then
               If yyy > wjry - 30 Then
                    If yyy < wjry + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
                    End If
                End If
         End If
    End If

     If xxx > wjraltkdlfx2 - 30 Then
          If xxx < wjraltkdlfx2 + 30 Then
               If yyy > wjraltkdlfy2 - 30 Then
                    If yyy < wjraltkdlfy2 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
wjraltkdlfx2 = wjrx
wjraltkdlfy2 = wjry
                    End If
                End If
         End If
    End If

          If xxx > wjraltkdlfx3 - 30 Then
          If xxx < wjraltkdlfx3 + 30 Then
               If yyy > wjraltkdlfy3 - 30 Then
                    If yyy < wjraltkdlfy3 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
wjraltkdlfx3 = wjrx
wjraltkdlfy3 = wjry
                    End If
                End If
         End If
    End If

     If xxx > wjraltkdlfx4 - 30 Then
          If xxx < wjraltkdlfx4 + 30 Then
               If yyy > wjraltkdlfy4 - 30 Then
                    If yyy < wjraltkdlfy4 + 30 Then

vhrvkfx4 = xxx
vhrvkfy4 = yyy
Call timer99
ahrt = ahrt - 1
wjraltkdlfx4 = wjrx
wjraltkdlfy4 = wjry
                    End If
                End If
         End If
    End If


     
     
     If wjrx > altkdlfx1 - 30 Then
          If wjrx < altkdlfx1 + 30 Then
               If wjry > altkdlfy1 - 30 Then
                    If wjry < altkdlfy1 + 30 Then

vhrvkfx4 = wjrx
vhrvkfy4 = wjry
Call timer99
altkdlfx1 = -30
wjr = wjr - 1
wjatn = wjatn + 3
                    
                    End If
                End If
         End If
    End If


     If xxx > vlftkfrlx - 30 Then
          If xxx < vlftkfrlx + 30 Then
               If yyy > vlftkfrly - 30 Then
                    If yyy < vlftkfrly + 30 Then

vlftkfrl1 = 0
vlftkfrly = -500
Image1.Visible = True
                    End If
                End If
         End If
    End If

End Sub

Private Sub tkdghkd()
Label1.Caption = "점수 : " & wjatn
Label2.Caption = "몫 : " & ahrt
End Sub

Private Sub altkdlf()

If sss = 1 Then
    If ssss = 0 Then

ssss = 1
altkdlfx1 = xxx
altkdlfy1 = yyy
End If
    End If
If altkdlfy1 < -10 Then
ssss = 0
sss = 0
End If

End Sub

Private Sub rlfhr()
     ReDim PLAYER(1 To 11) As String
     ReDim SCORE(1 To 11) As Integer
     Dim PLAYERNAME As String
     Dim PLAYERTEMP As String
     Dim SCORETEMP As Integer
     Dim MAINLOOP As Integer
     Dim LOOPCTR As Integer
     Dim FOUNDSW As Integer
     Open App.Path & "\game.txt" For Input As #1
     For LOOPCTR = 1 To 10
          Input #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     For LOOPCTR = 1 To 10
          If wjatn > SCORE(LOOPCTR) Then FOUNDSW = 1
     Next LOOPCTR
     If FOUNDSW = 1 Then
          PLAYERNAME = InputBox$("이름을 입력하세요")
          If PLAYERNAME = "" Then PLAYERNAME = "이름없음"
          PLAYER(11) = PLAYERNAME
          SCORE(11) = wjatn
     End If
     For MAINLOOPCTR = 1 To 10
          For LOOPCTR = 1 To 10
               If SCORE(LOOPCTR) < SCORE(LOOPCTR + 1) Then
                    PLAYERTEMP = PLAYER(LOOPCTR)
                    PLAYER(LOOPCTR) = PLAYER(LOOPCTR + 1)
                    PLAYER(LOOPCTR + 1) = PLAYERTEMP
                    SCORETEMP = SCORE(LOOPCTR)
                    SCORE(LOOPCTR) = SCORE(LOOPCTR + 1)
                    SCORE(LOOPCTR + 1) = SCORETEMP
               End If
          Next LOOPCTR
     Next MAINLOOPCTR
     Open "game.txt" For Output As #1
     MESSAGE = "높은 점수" + Chr$(13)
     For LOOPCTR = 1 To 10
          MESSAGE = MESSAGE + Chr$(13)
          If SCORE(LOOPCTR) = SCORECTR Then
               MESSAGE = MESSAGE + "-> " + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          Else
               MESSAGE = MESSAGE + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          End If
          Write #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     MsgBox (MESSAGE)

End Sub

Private Sub upwjraltkdlf1()
     upwjraltkdlfy1 = upwjraltkdlfy1 + 10
     If upwjraltkdlfy1 > 700 Then
          upwjraltkdlfx1 = upwjrx1
          upwjraltkdlfy1 = upwjry1
     End If
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx1, upwjraltkdlfy1, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx1, upwjraltkdlfy1, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     upwjraltkdlfy2 = upwjraltkdlfy2 + 15
     If upwjraltkdlfy2 > 700 Then
          upwjraltkdlfx2 = upwjrx2
          upwjraltkdlfy2 = upwjry2
     End If
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx2, upwjraltkdlfy2, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx2, upwjraltkdlfy2, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     upwjraltkdlfy3 = upwjraltkdlfy3 + 11
     If upwjraltkdlfy3 > 700 Then
          upwjraltkdlfx3 = upwjrx3
          upwjraltkdlfy3 = upwjry3
     End If
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx3, upwjraltkdlfy3, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx3, upwjraltkdlfy3, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     upwjraltkdlfy4 = upwjraltkdlfy4 + 11
     If upwjraltkdlfy4 > 700 Then
          upwjraltkdlfx4 = upwjrx4
          upwjraltkdlfy4 = upwjry4
     End If
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx4, upwjraltkdlfy4, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx4, upwjraltkdlfy4, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     upwjraltkdlfy5 = upwjraltkdlfy5 + 9
     If upwjraltkdlfy5 > 700 Then
          upwjraltkdlfx5 = upwjrx5
          upwjraltkdlfy5 = upwjry5
     End If
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx5, upwjraltkdlfy5, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, upwjraltkdlfx5, upwjraltkdlfy5, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub upwjraltkdlf2()
     upwjraltkdlfy3 = upwjraltkdlfy3 + 9
     If upwjraltkdlfy3 > 700 Then
          upwjraltkdlfx3 = upwjrx3
          upwjraltkdlfy3 = upwjry3
     End If

     upwjraltkdlfy4 = upwjraltkdlfy4 + 9
     If upwjraltkdlfy4 > 700 Then
          upwjraltkdlfx4 = upwjrx4
          upwjraltkdlfy4 = upwjry4
End If
End Sub
'Private Sub leftwjraltkdlf()
'     leftwjraltkdlfy1 = leftwjraltkdlfy1 + 8
'     If leftwjraltkdlfx1 > 700 Then
'          leftwjraltkdlfx1 = leftwjrx1
'          leftwjraltkdlfy1 = leftwjry1
'     End If
'     SUCCESS = BitBlt(Form1.hDC, leftwjraltkdlfx1, leftwjraltkdlfy1, Form2.Picture10.ScaleWidth, Form2.Picture10.ScaleHeight, Form2.Picture10.hDC, 0, 0, SRCAND)
'     SUCCESS = BitBlt(Form1.hDC, leftwjraltkdlfx1, leftwjraltkdlfy1, Form2.Picture9.ScaleWidth, Form2.Picture9.ScaleHeight, Form2.Picture9.hDC, 0, 0, SRCPAINT)

'End Sub

Private Sub startime()
If starttime > 0 And starttime < 8000 Then
Call upwjr1
Call cndehf
Call tkdghkd
Call altkdlf
Call upwjraltkdlf1
End If
If starttime > 200 And starttime < 800 Then
Call upwjr2
Call upwjraltkdlf2
End If
If starttime > 400 And starttime < 800 Then
Call leftwjr
Form2.Picture5.Picture = LoadPicture(App.Path & "\anti_right1.bmp")
Form2.Picture6.Picture = LoadPicture(App.Path & "\anti_right2.bmp")

End If
If starttime > 600 And starttime < 800 Then
Call rightwjr
Form2.Picture7.Picture = LoadPicture(App.Path & "\anti_left1.bmp")
Form2.Picture8.Picture = LoadPicture(App.Path & "\anti_left2.bmp")

End If

If starttime = 800 Then
MsgBox "다른 곳로 갑니다...2 스테이지"
'midi1.Command = "colse"
'midi1.FileName = App.Path + "\카탈로그.mid"
'midi1.Command = "open"
'midi1.Command = "play"

Label3.Caption = "2 스테이지"
Call dlxks
Form2.Picture3.Picture = LoadPicture(App.Path & "\적비행기흑.bmp")
Form2.Picture4.Picture = LoadPicture(App.Path & "\적비행기백.bmp")
wjatn = wjatn + 100
End If
If ahrt <= 0 Then
Timer1.Enabled = False
MsgBox "game over"
Call rlfhr
End
End If

End Sub

Private Sub dlxks()
     Randomize Timer
     Call starts
BACKGROUNDCTR = 430
     Form3.Picture1.Picture = LoadPicture(App.Path & "\back.BMP")
'     Picture = LoadPicture(App.Path & "\배경그림.bmp")
  
 
 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture2.ScaleWidth, Form2.Picture2.ScaleHeight, Form2.Picture2.hDC, 0, 0, SRCAND)
 SUCCESS = BitBlt(Form1.hDC, xxx, yyy, Form2.Picture1.ScaleWidth, Form2.Picture1.ScaleHeight, Form2.Picture1.hDC, 0, 0, SRCPAINT)



End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Timer3_Timer()
End Sub
Private Sub king()
upwjry1 = -10
upwjrx1 = 0
upwjrcnt1 = 0
upwjry2 = -10
upwjrx2 = 0
upwjrcnt2 = 0
upwjry3 = -10
upwjrx3 = 0
upwjrcnt3 = 0
upwjry4 = -10
upwjrx4 = 0
upwjrcnt4 = 0
upwjry5 = -10
upwjrx5 = 0
upwjrcnt5 = 0

upwjrspeed1 = 5
upwjrspeed2 = 5
upwjrspeed3 = 5
upwjrspeed4 = 5
upwjrspeed5 = 5

leftwjry1 = 200
leftwjrx1 = -50
leftwjrcnt1 = 0
leftwjrspeed1 = 5
leftwjry2 = 200
leftwjrx2 = -50
leftwjrcnt2 = 0
leftwjrspeed2 = 5

rightwjry1 = 200
rightwjrx1 = 600
rightwjrcnt1 = 0
rightwjrspeed1 = 5

rightwjry2 = 300
rightwjrx2 = 700
rightwjrcnt2 = 0
rightwjrspeed2 = 5

altkdlfx1 = -50
altkdlfy1 = -50

upwjraltkdlfx1 = 1000
upwjraltkdlfy1 = 0
upwjraltkdlfx2 = 1000
upwjraltkdlfy2 = 0
upwjraltkdlfx3 = 1000
upwjraltkdlfy3 = 0
upwjraltkdlfx4 = 1000
upwjraltkdlfy4 = 0
upwjraltkdlfx5 = 1000
upwjraltkdlfy5 = 0

wjry = 50
wjrx = 200
wjrcnt = 1
wjr = 10
End Sub

Private Sub vlftkfrl()
If Image1.Visible = False Then
vlftkfrl1 = vlftkfrl1 + 1
If vlftkfrl1 = Int(Rnd(20) * 50) Then
     Randomize Timer
vlftkfrlx = Int(Rnd(100) * 300)
vlftkfrl1 = 0
End If
End If
vlftkfrly = vlftkfrly + 8
     
     SUCCESS = BitBlt(Form1.hDC, vlftkfrlx, vlftkfrly, Form2.Picture12.ScaleWidth, Form2.Picture12.ScaleHeight, Form2.Picture12.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vlftkfrlx, vlftkfrly, Form2.Picture11.ScaleWidth, Form2.Picture11.ScaleHeight, Form2.Picture11.hDC, 0, 0, SRCPAINT)
If vlftkfrly > 1000 Then
     Randomize Timer
vlftkfrly = -500
End If

End Sub

Private Sub timer33()
Label4.Visible = True
Label4.Caption = "왕의 에너지 : " & wjr
Label3.Caption = "왕이다"
Call cndehf
Call altkdlf
starttime = starttime + 1
 
     BACKGROUNDCTR = BACKGROUNDCTR - 1
     If BACKGROUNDCTR = 0 Then
     BACKGROUNDCTR = 450
         End If
     SUCCESS = BitBlt(Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form3.Picture1.hDC, 0, BACKGROUNDCTR, SRCCOPY)
Call play


If xxx < 0 Then
xxx = 0
ElseIf yyy < 0 Then
yyy = 0
ElseIf yyy > 430 Then
yyy = 430
ElseIf xxx > 560 Then
xxx = 560
End If

If wjrx < 0 Then wjrcnt = wjrcnt + 1
If wjrx > 560 Then wjrcnt = wjrcnt + 1
If wjrcnt = 1 Then wjrx = wjrx + 5
If wjrcnt = 2 Then wjrx = wjrx - 5
If wjrcnt = 3 Then wjrx = wjrx + 5
If wjrcnt = 4 Then wjrx = wjrx - 5
If wjrcnt = 5 Then wjrx = wjrx + 5
If wjrcnt = 6 Then wjrx = wjrx - 5
If wjrcnt = 7 Then wjrx = wjrx + 5
If wjrcnt = 8 Then wjrx = wjrx - 5
If wjrcnt = 9 Then wjrx = wjrx + 5
If wjrcnt = 10 Then wjrx = wjrx - 5
If wjrcnt = 11 Then wjrx = wjrx + 5
If wjrcnt = 12 Then wjrx = wjrx - 5
If wjrcnt = 13 Then wjrx = wjrx + 5
If wjrcnt = 14 Then wjrx = wjrx - 5



SUCCESS = BitBlt(Form1.hDC, wjrx, wjry, Form2.Picture4.ScaleWidth, Form2.Picture4.ScaleHeight, Form2.Picture4.hDC, 0, 0, SRCAND)
SUCCESS = BitBlt(Form1.hDC, wjrx, wjry, Form2.Picture3.ScaleWidth, Form2.Picture3.ScaleHeight, Form2.Picture3.hDC, 0, 0, SRCPAINT)
If BACKGROUNDCTR = 1 Then
MsgBox "당신은 시간이 지나 패배하였습니다."
Call rlfhr
End
End If

     wjraltkdlfy = wjraltkdlfy + 20
     If wjraltkdlfy > 500 Then
          wjraltkdlfx = wjrx
          wjraltkdlfy = wjry
     End If
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx, wjraltkdlfy, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx, wjraltkdlfy, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     wjraltkdlfy2 = wjraltkdlfy2 + 18
     If wjraltkdlfy2 > 500 Then
          wjraltkdlfx2 = wjrx
          wjraltkdlfy2 = wjry
     End If
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx2, wjraltkdlfy2, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx2, wjraltkdlfy2, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     wjraltkdlfy3 = wjraltkdlfy3 + 16
     If wjraltkdlfy3 > 500 Then
          wjraltkdlfx3 = wjrx
          wjraltkdlfy3 = wjry
     End If
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx3, wjraltkdlfy3, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx3, wjraltkdlfy3, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

     wjraltkdlfy4 = wjraltkdlfy4 + 14
     If wjraltkdlfy4 > 500 Then
          wjraltkdlfx4 = wjrx
          wjraltkdlfy4 = wjry
     End If
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx4, wjraltkdlfy4, Form2.Picture18.ScaleWidth, Form2.Picture18.ScaleHeight, Form2.Picture18.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, wjraltkdlfx4, wjraltkdlfy4, Form2.Picture17.ScaleWidth, Form2.Picture17.ScaleHeight, Form2.Picture17.hDC, 0, 0, SRCPAINT)

If wjr = 0 Then
MsgBox "당신을 승리하였습니다."
wjatn = wjatn + 1000
Call rlfhr
End
End If
Label1.Caption = "점수 : " & wjatn
Label2.Caption = "몫 : " & ahrt
If ahrt <= 0 Then
MsgBox "당신은 패배하였습니다."
Call rlfhr
End
End If


End Sub

Private Sub timer22()


Label3.Caption = "2 스테이지"
starttime = starttime + 1
 
     BACKGROUNDCTR = BACKGROUNDCTR - 1
     If BACKGROUNDCTR = 0 Then
     BACKGROUNDCTR = 450
         End If
     SUCCESS = BitBlt(Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form3.Picture1.hDC, 0, BACKGROUNDCTR, SRCCOPY)
Call play


If xxx < 0 Then
xxx = 0
ElseIf yyy < 0 Then
yyy = 0
ElseIf yyy > 430 Then
yyy = 430
ElseIf xxx > 560 Then
xxx = 560
End If
Call startime
If BACKGROUNDCTR = 1 Then
BACKGROUNDCTR = 430
Call king
MsgBox "왕이다!."
Form2.Picture3.Picture = LoadPicture(App.Path & "\boss1.bmp")
Form2.Picture4.Picture = LoadPicture(App.Path & "\boss2.bmp")

'midi1.FileName = App.Path + "\카탈로그.mid"
'midi1.Command = "open"
'midi1.Command = "play"

Label3.Caption = "왕이다"
BACKGROUNDCTR = 430
End If
Label1.Caption = "점수 : " & wjatn
Label2.Caption = "몫 : " & ahrt
If Image1.Visible = False Then Call vlftkfrl

End Sub
Private Sub timer11()
Label3.Caption = "1 스테이지"
starttime = starttime + 1
     BACKGROUNDCTR = BACKGROUNDCTR - 1
     If BACKGROUNDCTR = 0 Then
     BACKGROUNDCTR = 450
         End If
     SUCCESS = BitBlt(Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form3.Picture1.hDC, 0, BACKGROUNDCTR, SRCCOPY)
'If BACKGROUNDCTR >= 1000 Then
'Timer2.Enabled = True
'Timer1.Enabled = False
'Else
'Timer1.Enabled = True
'End If
'     If BACKGROUNDCTR = 10 Then
'    BACKGROUNDCTR = -500
'    SUCCESS = BitBlt(Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form3.Picture1.hDC, 0, BACKGROUNDCTRa, SRCCOPY)
'     End If
Call play


If xxx < 0 Then
xxx = 0
ElseIf yyy < 0 Then
yyy = 0
ElseIf yyy > 430 Then
yyy = 430
ElseIf xxx > 560 Then
xxx = 560
End If
Call startime
If Image1.Visible = False Then Call vlftkfrl

End Sub
Private Sub timer66()
Static c As Integer
c = c + 1
If c = 9 Then
c = 1
End If
Form2.Picture13.Picture = ImageList1.ListImages(c).Picture
Form2.Picture14.Picture = ImageList2.ListImages(c).Picture
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx1, vhrvkfy1, Form2.Picture14.ScaleWidth, Form2.Picture14.ScaleHeight, Form2.Picture14.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx1, vhrvkfy1, Form2.Picture13.ScaleWidth, Form2.Picture13.ScaleHeight, Form2.Picture13.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub timer77()
Static b As Integer
b = b + 1
If b = 9 Then
b = 1
End If
Form2.Picture13.Picture = ImageList1.ListImages(b).Picture
Form2.Picture14.Picture = ImageList2.ListImages(b).Picture
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx2, vhrvkfy2, Form2.Picture14.ScaleWidth, Form2.Picture14.ScaleHeight, Form2.Picture14.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx2, vhrvkfy2, Form2.Picture13.ScaleWidth, Form2.Picture13.ScaleHeight, Form2.Picture13.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub timer88()
Static a As Integer
a = a + 1
If a = 9 Then
a = 1
End If
Form2.Picture13.Picture = ImageList1.ListImages(a).Picture
Form2.Picture14.Picture = ImageList2.ListImages(a).Picture
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx3, vhrvkfy3, Form2.Picture14.ScaleWidth, Form2.Picture14.ScaleHeight, Form2.Picture14.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx3, vhrvkfy3, Form2.Picture13.ScaleWidth, Form2.Picture13.ScaleHeight, Form2.Picture13.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub timer99()
Static d As Integer
d = d + 1
If d = 9 Then
d = 1
End If
Form2.Picture13.Picture = ImageList1.ListImages(d).Picture
Form2.Picture14.Picture = ImageList2.ListImages(d).Picture
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx4, vhrvkfy4, Form2.Picture14.ScaleWidth, Form2.Picture14.ScaleHeight, Form2.Picture14.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx4, vhrvkfy4, Form2.Picture13.ScaleWidth, Form2.Picture13.ScaleHeight, Form2.Picture13.hDC, 0, 0, SRCPAINT)

End Sub

Private Sub timer100()
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx5, vhrvkfy5, Form2.Picture16.ScaleWidth, Form2.Picture16.ScaleHeight, Form2.Picture16.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(Form1.hDC, vhrvkfx5, vhrvkfy5, Form2.Picture15.ScaleWidth, Form2.Picture15.ScaleHeight, Form2.Picture15.hDC, 0, 0, SRCPAINT)

End Sub

