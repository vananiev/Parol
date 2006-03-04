VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Infomation"
   ClientHeight    =   1365
   ClientLeft      =   4470
   ClientTop       =   6240
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Help > Vitalij"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rgs As New REGISTRY
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndlnsertAfter As Long, ByVal x As Long, _
ByVal ó As Long, _
ByVal ex As Long, _
ByVal cy As Long, _
ByVal wFlags As Long _
) As Long
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPHOST = -1
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPHOST = -2
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Form_Load()
Rgs.AppName = "Stop"
Rgs.Section = Form1.txtName
Rgs.Key = "HelpParol"
Rgs.RegGet
Label1 = Rgs.Setting

End Sub
