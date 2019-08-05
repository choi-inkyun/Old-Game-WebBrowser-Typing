Attribute VB_Name = "Module1"
Const snd_async = &H1
Const snd_syne = &H0
Const snd_loop = &H8
Const snd_memory = &H4
Const snd_nodefault = &H2
Const snd_nostop = &H10
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Global wjatn, ahrt As Integer
Global SUCCES As Integer

Global Const SRCCOPY = &HCC0020
Global Const SRCAND = &H8800C6
Global Const SRCPAINT = &HEE0086
Global xxx As Integer
Global yyy As Integer
Global lefts As Integer
Global rights As Integer
Global ups As Integer
Global downs As Integer
Global BACKGROUNDCTR As Integer
Global BACKGROUNDCTRA As Integer
Global starttime As Integer
Global upwjrx1 As Integer
Global upwjry1 As Integer
Global upwjrcnt1 As Integer
Global upwjrspeed1 As Integer
Global upwjrx2 As Integer
Global upwjry2 As Integer
Global upwjrcnt2 As Integer
Global upwjrspeed2 As Integer
Global upwjrx3 As Integer
Global upwjry3 As Integer
Global upwjrcnt3 As Integer
Global upwjrspeed3 As Integer
Global upwjrx4 As Integer
Global upwjry4 As Integer
Global upwjrcnt4 As Integer
Global upwjrspeed4 As Integer
Global upwjrx5 As Integer
Global upwjry5 As Integer
Global upwjrcnt5 As Integer
Global upwjrspeed5 As Integer
Global leftwjry1 As Integer
Global leftwjrx1 As Integer
Global leftwjrcnt1 As Integer
Global leftwjrspeed1 As Integer
Global leftwjry2 As Integer
Global leftwjrx2 As Integer
Global leftwjrcnt2 As Integer
Global leftwjrspeed2 As Integer
Global rightwjry1 As Integer
Global rightwjrx1 As Integer
Global rightwjrcnt1 As Integer
Global rightwjrspeed1 As Integer
Global rightwjry2 As Integer
Global rightwjrx2 As Integer
Global rightwjrcnt2 As Integer
Global rightwjrspeed2 As Integer

Global altkdlfx1 As Integer
Global altkdlfy1 As Integer
Global sss As Integer
Global ssss As Integer
Global upwjraltkdlfx1 As Integer
Global upwjraltkdlfy1 As Integer
Global upwjraltkdlfx2 As Integer
Global upwjraltkdlfy2 As Integer
Global upwjraltkdlfx3 As Integer
Global upwjraltkdlfy3 As Integer
Global upwjraltkdlfx4 As Integer
Global upwjraltkdlfy4 As Integer
Global upwjraltkdlfx5 As Integer
Global upwjraltkdlfy5 As Integer

Global wjry As Integer
Global wjrx As Integer
Global wjrcnt As Integer

Global wjraltkdlfx As Integer
Global wjraltkdlfy As Integer
Global wjraltkdlfx2 As Integer
Global wjraltkdlfy2 As Integer
Global wjraltkdlfx3 As Integer
Global wjraltkdlfy3 As Integer
Global wjraltkdlfx4 As Integer
Global wjraltkdlfy4 As Integer

Global wjr As Integer
Global vlftkfrlx As Integer
Global vlftkfrly As Integer
Global vhrvkfy1 As Integer
Global vhrvkfx1 As Integer
Global vhrvkfy2 As Integer
Global vhrvkfx2 As Integer
Global vhrvkfy3 As Integer
Global vhrvkfx3 As Integer
Global vhrvkfy4 As Integer
Global vhrvkfx4 As Integer
Global vhrvkfy5 As Integer
Global vhrvkfx5 As Integer

Global vlftkfrl1 As Integer

