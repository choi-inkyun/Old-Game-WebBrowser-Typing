VERSION 5.00
Object = "{E54B6DC3-AE1F-11D1-A750-006097310C00}#1.0#0"; "GIFPLAY.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  '����
   Caption         =   "Form2"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   3585
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   0
   End
   Begin GIFPLAYLib.GifPlay GifPlay1 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   8281
      _StockProps     =   161
      AnimationGifFileName=   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�������� ����ϴ��� �ý����� �ټ� ���������� ������ �������ֽñ� �ٶ��ϴ�.         <-- Ŭ��"
      Height          =   1935
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub GifPlay1_Click()
MsgBox "���α�..��3 ������ id : dingpong, E-Mail : dingpong@hitel.net, ���α׷����� ����ϰ�����." + Chr(13) + Chr(13) _
       + "������..��2 ������ id : Garshion, E-Mail : Garshion@hitel.net, �׷��� ���� �����."
End Sub

Private Sub Timer1_Timer()
  If Timer1.Enabled = True Then
     
    Call GifPlay1.LoadAnimationGifFile(App.Path & "\Maximum.gif")
            
    If GifPlay1.Play = False Then
      
     Else
     
     End If
      
     Timer1.Enabled = False
     
  End If
End Sub
