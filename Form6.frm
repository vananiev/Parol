VERSION 5.00
Begin VB.Form frmSFG 
   Caption         =   "Safeguard Sistem"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   360
   End
End
Attribute VB_Name = "frmSFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Form_Terminate()
    Unload frmSFG
End Sub

Private Sub tmrWait_Timer()
    DoEvents
    If vSafeguard = "Exit" Then Unload frmSFG
    If vSafeguard = "Error" Then
        ExitWindowsEx 4, 0
    Else
        vSafeguard = "Error"
    End If
End Sub

