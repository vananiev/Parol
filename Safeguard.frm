VERSION 5.00
Begin VB.Form frmSafeguard 
   Caption         =   "Safeguard Sistem"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrWait 
      Interval        =   50
      Left            =   1680
      Top             =   240
   End
End
Attribute VB_Name = "frmSafeguard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strStatus As String
Dim Object As New Safeguard.clsSFG
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Form_Load()
    strStatus = "All_right"
End Sub

Private Sub tmrWait_Timer()
     On Error GoTo Shell
     Object.SFG = strStatus
     Exit Sub
Shell:
    ExitWindowsEx 4, 0
End Sub

