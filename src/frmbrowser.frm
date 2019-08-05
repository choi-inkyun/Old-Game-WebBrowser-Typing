VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmbrowser 
   Caption         =   "MaximumSoft Nice Internet Browser Test Version"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox picaddress 
      Height          =   15
      Left            =   5400
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   360
      Width           =   15
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '아래 맞춤
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   6015
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   11456
            MinWidth        =   11465
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2001-05-15"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 4:08"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlicons 
      Left            =   8760
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":36A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":51FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":6D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbrowser.frx":88A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   8880
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   2160
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   4695
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   9495
      ExtentX         =   16748
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox Combo 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   8655
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  '위 맞춤
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "imlicons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lbladdress 
      Caption         =   "주소 :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   615
   End
   Begin VB.Menu mnufile 
      Caption         =   "파일 (&F)"
      Begin VB.Menu mnunew 
         Caption         =   "새로운 창 (&N)"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "열기 (&O)"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "닫기 (&C)"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "종료하기 (&X)"
      End
   End
End
Attribute VB_Name = "frmbrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public startingaddress As String
Dim mbdontnavigatenow As Boolean



Private Sub Combo_Click()
If mbdontnavigatenow Then Exit Sub
Timer.Enabled = True
WebBrowser.Navigate Combo.Text

End Sub

Private Sub Combo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
Combo_Click
End If
End Sub

Private Sub Form_Load()
startingaddress = "http://www.daum.net"
On Error Resume Next
Me.Show
Toolbar.Refresh

If Len(startingaddress) > 0 Then
Combo.Text = startingaddress
Combo.AddItem Combo.Text
Timer.Enabled = True
WebBrowser.Navigate startingaddress
End If

End Sub

Private Sub Form_Resize()
Combo.Width = Me.ScaleWidth - 750
WebBrowser.Width = Me.ScaleWidth - 50
WebBrowser.Height = Me.ScaleHeight - (picaddress.Top + picaddress.Height + StatusBar.Height) - 1000
End Sub

Private Sub mnuclose_Click()
Me.Hide

End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnunew_Click()
Dim frmb As New frmbrowser
frmb.startingaddress = "http://www.daum.net"
frmb.Show
End Sub

Private Sub mnuopen_Click()
Dim sfile As String
With CommonDialog
.Filter = "HTML Files(*.htm)|*.htm"
.ShowOpen
If Len(.FileName) = 0 Then
Exit Sub
End If
sfile = .FileName
End With
WebBrowser.Navigate "file://" + CommonDialog.FileName
End Sub

Private Sub Timer_Timer()
If WebBrowser.Busy = False Then
Timer.Enabled = False
StatusBar.Panels(1).Text = WebBrowser.LocationName
Else
StatusBar.Panels(1).Text = "읽는중..."
End If
End Sub

Private Sub WebBrowser_DownloadComplete()
On Error Resume Next
StatusBar.Panels(1).Text = WebBrowser.LocationName
End Sub

Private Sub WebBrowser_NavigateComplete(ByVal pDisp As String)
Dim i As Integer
Dim bfound As Boolean
StatusBar.Panels(1).Text = WebBrowser.LocationName
For i = 0 To Combo.ListCount - 1
If Combo.List(i) = WebBrowser.LocationURL Then
bfound = True
Exit For
End If
Next i
mbdontnavigatenow = True
If bfound Then
Combo.RemoveItem i
End If
Combo.AddItem WebBrowser.LocationURL, 0
Combo.ListIndex = 0
mbdontnavigatenow = False
End Sub

Private Sub toolbar_buttonclick(ByVal button As button)
On Error Resume Next
Timer.Enabled = True
Select Case button.Key
Case "back"
WebBrowser.GoBack
Case "forwand"
WebBrowser.GoForward
Case "refresh"
WebBrowser.Refresh
Case "home"
WebBrowser.GoHome
Case "search"
WebBrowser.GoSearch
Case "stop"
Timer.Enabled = False
WebBrowser.Stop
StatusBar.Panels(1).Text = Combo.Text
End Select

End Sub
