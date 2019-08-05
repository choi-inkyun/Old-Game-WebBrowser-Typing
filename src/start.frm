VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   Picture         =   "start.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      Silent          =   -1  'True
      DeviceType      =   ""
      FileName        =   "C:\Program Files\Microsoft Visual Studio\VB98\start.mid"
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Click Picture!"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Form1.Hide

End Sub

Private Sub Form_Load()

MMControl1.Command = "open"
MMControl1.Command = "play"
frmbrowser.Show

End Sub
