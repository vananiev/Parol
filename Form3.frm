VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Sistem not stabil"
   ClientHeight    =   1260
   ClientLeft      =   4890
   ClientTop       =   5310
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPrl 
      Interval        =   10
      Left            =   4080
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Система не стабильна"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public btLife As Integer
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Command1_Click()
    ExitWindowsEx 4, 0
End Sub

Private Sub tmrPrl_Timer()
    btLife = btLife + 1
    If btLife >= 300 Then ExitWindowsEx 4, 0
End Sub
