VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "�ִϰ� ��ǻ�� ���Ӱ� ������ ���α� ���� ���α׷�"
   ClientHeight    =   1410
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6420
   Icon            =   "�ִϰ�...frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "�ִϰ�...frx":030A
   ScaleHeight     =   1410
   ScaleWidth      =   6420
   StartUpPosition =   2  'ȭ�� ���
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "�ִϰ�...frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "�ִϰ�...frx":0930
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "�ִϰ�...frx":0C4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbxld"
            Object.ToolTipText     =   "����"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "xkwk"
            Object.ToolTipText     =   "Ÿ��"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nib"
            Object.ToolTipText     =   "��������"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   480
      MaxLength       =   28
      TabIndex        =   2
      Text            =   "�ִϰ� ��ǻ�� ���Ӱ� ������ ���α� ���� ���α׷�"
      ToolTipText     =   "��縦 �Է��մϴ�"
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Timer �ö󰡶� 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4440
      Top             =   600
   End
   Begin VB.Timer ���� 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   600
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "15"
      ToolTipText     =   "��簡 ���ܳ����� �ӵ��� �Է��մϴ�"
      Top             =   4830
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "15"
      ToolTipText     =   "��簡 ��������� �ӵ��� �Է��մϴ�"
      Top             =   4830
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Copyright �� 2ooo White LEE (E-mail : Lostmage@Chollian.net)"
      ForeColor       =   &H00008080&
      Height          =   180
      Left            =   15
      TabIndex        =   12
      Top             =   5295
      Width           =   5415
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӵ� 2"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   4455
      TabIndex        =   11
      Top             =   4890
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӵ� 1"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   3255
      TabIndex        =   10
      Top             =   4890
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   15
      TabIndex        =   9
      Top             =   4890
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   615
      Left            =   0
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400000&
      Height          =   615
      Left            =   0
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label z12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   0
      TabIndex        =   8
      Top             =   4860
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Height          =   180
      Left            =   3120
      TabIndex        =   6
      Top             =   4800
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӵ� 1"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3240
      TabIndex        =   5
      Top             =   4860
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ӵ� 2"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   4440
      TabIndex        =   4
      Top             =   4860
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   5400
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Copyright �� 2ooo White LEE (E-mail : Lostmage@Chollian.net)"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Menu file 
      Caption         =   "����(&f)"
      Begin VB.Menu end 
         Caption         =   "����"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu ghjfj 
      Caption         =   "����(&g)"
      Begin VB.Menu tbxld 
         Caption         =   "����"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu ghfh 
      Caption         =   "��ƿ��Ƽ(&u)"
      Begin VB.Menu muxkwk 
         Caption         =   "Ÿ��"
         Shortcut        =   ^T
      End
      Begin VB.Menu web 
         Caption         =   "�� ������"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu wpwkr 
      Caption         =   "����(&j)"
      Begin VB.Menu mutame 
         Caption         =   "����"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ����(100)
Dim ����� As Boolean
Dim ���簳��
Dim ���
Private Sub ��ξ�����()
Dim I As Integer
For I = 1 To ���簳��
Unload Label1(I)
Next I
Timer1.Enabled = True
End Sub
Private Sub ����_Timer()
Static A As Integer
Dim ǥ��
If A = Len(Text1.Text) Then ����.Enabled = False: A = 0: GoTo 1
A = A + 1
ǥ�� = Mid(���, A, 1)
Load Label1(A)
Label1(A).Move Label1(A - 1).Left + Label1(A - 1).Width, Label1(A - 1).Top
Label1(A).ForeColor = RGB(0, 0, 0)
Label1(A).Visible = True
Label1(A).Caption = ǥ��
Label1(A).ForeColor = RGB(0, 0, 0)
�ö󰡶�.Enabled = True
���簳�� = A
1 End Sub

Private Sub �ö󰡶�_Timer()
Static �Ϸᰳ��
Static b As Integer
Dim I As Integer
If ����� = True Then GoTo 2 Else GoTo 1
1 For I = 0 To ���簳��
����(I) = ����(I) + Val(Text2.Text)
Label1(I).ForeColor = RGB(����(I), ����(I), ����(I))
If ����(I) >= 255 Then ����(I) = 255: �Ϸᰳ�� = I
Next I
GoTo 3
2 For I = 0 To ���簳��
����(I) = ����(I) - Val(Text3.Text)
If ����(I) <= 0 Then ����(I) = 0: �Ϸᰳ�� = I
Label1(I).ForeColor = RGB(����(I), ����(I), ����(I))
Next I
GoTo 4
3 If �Ϸᰳ�� = Len(���) Then ����� = True: �Ϸᰳ�� = 0: GoTo 5
4 If �Ϸᰳ�� = Len(���) Then �ö󰡶�.Enabled = False: �Ϸᰳ�� = 0: Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True: ��ξ�����: ����� = False: ���簳�� = 0: GoTo 5
5 End Sub

Private Sub end_Click()
End
End Sub

Private Sub mutame_Click()
Form2.Show
End Sub

Private Sub muxkwk_Click()
Call Shell(App.Path & "\�α����� Ÿ�ڿ���.exe", 1)
End Sub

Private Sub tbxld_Click()
Call Shell(App.Path & "\����.exe", 1)
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If Len(Text2.Text) = 0 Then MsgBox "�ӵ� 1�� ���� �־��ּ���.", vbExclamation, "����": GoTo 2
2  If Len(Text3.Text) = 0 Then MsgBox "�ӵ� 2�� ���� �־��ּ���.", vbExclamation, "����": GoTo 3
 ��� = Text1.Text
 ����.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
End If
3 End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If Len(Text2.Text) = 0 Then MsgBox "�ӵ� 1�� ���� �־��ּ���.", vbExclamation, "����": GoTo 2
2 If Len(Text3.Text) = 0 Then MsgBox "�ӵ� 2�� ���� �־��ּ���.", vbExclamation, "����": GoTo 3
 ��� = Text1.Text
 ����.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
End If
3 End Sub

Private Sub Timer1_Timer()
  If Len(Text2.Text) = 0 Then MsgBox "�ӵ� 1�� ���� �־��ּ���.", vbExclamation, "����": GoTo 2
2 If Len(Text3.Text) = 0 Then MsgBox "�ӵ� 2�� ���� �־��ּ���.", vbExclamation, "����": GoTo 3

 ��� = Text1.Text
 ����.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
Timer1.Enabled = False
3 End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "tbxld"
Call Shell(App.Path & "\����.exe", 1)
Case "xkwk"
Call Shell(App.Path & "\�α����� Ÿ�ڿ���.exe", 1)
Case "nib"
Call Shell(App.Path & "\Nice.exe", 1)
End Select
End Sub

Private Sub web_Click()
Call Shell(App.Path & "\Nice.exe", 1)
End Sub
