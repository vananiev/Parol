VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7995
   ClientLeft      =   930
   ClientTop       =   2085
   ClientWidth     =   10530
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form6"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Object As Object
'API-функция для получения копии всего экрана
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal IngX As Long, ByVal IngY As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal IngXSrc As Long, ByVal IngYSrc As Long, _
ByVal dwRop As Long _
) As Long
'API-функция для получения описателя дисплейного изображения
Private Declare Function GetDesktopWindow _
Lib "user32" () As Long
'API-функция для получения контекста устройства по описателю
Private Declare Function GetDC _
Lib "user32" ( _
ByVal hwnd As Long _
) As Long
'API-функция,освобождающая контекст устройства
Private Declare Function ReleaseDC _
Lib "user32" ( _
ByVal hwnd As Long, _
ByVal hdc As Long _
) As Long

'константы для некоторых API-функций
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Declare Function SystemParametersInfo _
Lib "user32" Alias "SystemParametersInfoA" ( _
ByVal uAction As Long, _
ByVal uParam As Long, _
ByRef IpvParam As Any, _
ByVal fuWinlni As Long _
) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndlnsertAfter As Long, ByVal x As Long, _
ByVal у As Long, _
ByVal ex As Long, _
ByVal cy As Long, _
ByVal wFlags As Long _
) As Long
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPHOST = -1
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPHOST = -2

Private Sub Form_Click()
    Dim A As New Safeguard
    Form1.SetFocus
    Form1.Timer1.Enabled = True
    Object.Timer = True
    A.frmSFG.tmrWait.Enabled = True
End Sub

Private Sub Form_DblClick()
    Form_Click
    Object.Timer = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Object.Timer = True
End Sub

Private Sub Form_Load()
Dim dwRop As Long
Dim hwndSrc As Long
Dim hSrcDC As Long
Dim lngRes As Long
Set Object = Form1
'каждый раз показываем разные изображения
Randomize
'копируем весь экран в окно рисунка
ScaleMode = vbPixels
Move 0, 0, Screen.Width + 1, Screen.Height + 1
dwRop = &HCC0020
hwndSrc = GetDesktopWindow()
hSrcDC = GetDC(hwndSrc)
lngRes = BitBlt(Form6.hdc, 0, 0, ScaleWidth, ScaleHeight, hSrcDC, 0, 0, dwRop)
lngRes = ReleaseDC(hwndSrc, hSrcDC)
SetWindowPos hwnd, HWND_TOPHOST, _
0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'показываем форму в полноэкранном режиме
Show
End Sub

