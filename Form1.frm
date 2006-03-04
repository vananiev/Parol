VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Controler"
   ClientHeight    =   2040
   ClientLeft      =   4590
   ClientTop       =   4530
   ClientWidth     =   4815
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   0
   End
   Begin VB.ListBox List 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrSleep 
      Interval        =   1000
      Left            =   720
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C9DFE0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C9DFE0&
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Введите имя:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "секунд"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "попыток"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Введите код:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim m As Integer
Dim blnOpen As Boolean
Dim Rgs As New REGISTRY
Dim sX, sY As Single
Dim lnSec As Long
    Dim bScroll() As Byte
    Dim sScroll() As String
    Dim intLastVal As Integer
    Dim bCurN As Integer
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

Private Sub Command1_Click()
If n <= 0 Then Form3.Show vbModal
If Text1.Text = "" Or txtName = "" Then GoTo 1
If Text1.Text = GetSetting("Stop", txtName, "Parol") Then
    frmSafeguard.strStatus = "Exit"
    blnOpen = True
    Timer1.Enabled = False
    ' отменяем у формы статус "самой верхней"
    SetWindowPos hwnd, HWND_NOTOPHOST, _
    0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Form6.Visible = False
    Form4.Show vbModal
End If
1:
n = n - 1
Label3 = n
If Text1.Text = "help" Then
    ' отменяем у формы статус "самой верхней"
    SetWindowPos hwnd, HWND_NOTOPHOST, _
    0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Form2.Show vbModal
    SetWindowPos hwnd, HWND_TOPHOST, _
    0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys vbTab
    KeyAscii = 0
End If
Object.Timer = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not (Timer1.Enabled) Then Timer1.Enabled = True
End Sub

Private Sub Form_Load()
If GetSetting("Stop", "Default", "Ferst", 0) = 1 Then
    Unload Form6
    Unload Form1
    frmSafeguard.strStatus = "Exit"
    MsgBox "Создайте своего пользователя", vbInformation, "Parol"
    DeleteSetting "Stop", "Default", "Ferst"
    Form5.Show vbModal
    End
End If
If GetSetting("Stop", "Default", "Start", 0) = 1 Then End
' Режим сетчатки
If GetSetting("Stop", txtName, "ScrollMode") <> "False" Then
    For n = 0 To 9
        List.AddItem Str(n)
    Next n
    List.SetFocus
    sScroll = Split(GetSetting("Stop", txtName, "ScrollMode"), "-")
    For n = 0 To UBound(sScroll) - 1
        bScroll(n) = Val(sScroll(n))
    Next n
    End
End If
n = 10: m = 60
Label2.Caption = "У вас осталось"
Label3.Caption = "10"
Label6 = "60"
blnOpen = False
Rgs.AppName = "Stop"
SetWindowPos hwnd, HWND_TOPHOST, _
0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_LostFocus()
    Timer1.Enabled = True
End Sub

Private Sub Form_Terminate()
    If Not (blnOpen) Then frmSafeguard.strStatus = "Error": Form3.Show
End Sub

Private Sub Timer1_Timer()
m = m - 1
Label6.Caption = Str(m)
If m <= -1 Then frmSafeguard.strStatus = "Error": Form3.Show vbModal
End Sub

Private Sub Timer2_Timer()
    lnSec = lnSec + 1
    If lnSec > 120 Then frmSafeguard.strStatus = "Error"
End Sub

Private Sub List_Scroll()
    If GetSetting("Stop", txtName, "ScrollMode") <> "False" Then tmrScr.Enabled = True
    intLastVal = -2
End Sub

Private Sub tmrScr_Timer()
    If intLastVal = List.List(List.ListIndex) Then
        If bScroll(bCurN) <> List.List(List.ListIndex) Then frmSafeguard.strStatus = "Error": Form3.Label1.Caption = "Ваша сечатка не входит в базу": Form3.Show vbModal
        If bCurN = UBound(sScroll) - 1 Then End
        List.ListIndex = 0
        bCurN = bCurN + 1
    End If
    intLastVal = List.List(List.ListIndex)
End Sub
