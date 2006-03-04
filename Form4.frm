VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1095
   ClientLeft      =   4920
   ClientTop       =   5175
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Setting"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub


Private Sub Command2_Click()
    Form5.Show vbModal
End Sub

Private Sub Form_Load()
    frmSafeguard.strStatus = "Exit"
    Form4.Caption = "Hello " & Form1.txtName
End Sub

Private Sub Form_Terminate()
    Unload frmSafeguard
    Unload Form1
    Unload Form4
End Sub
