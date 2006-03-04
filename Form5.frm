VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Setting"
   ClientHeight    =   2490
   ClientLeft      =   4785
   ClientTop       =   5565
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMode 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Удалить"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Check1"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   7
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtPrl 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtHelp 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "ScrollMode"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Не запускать при входе в систему"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblUser 
      Caption         =   "User="
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Parol="
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Help="
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rgs As New REGISTRY

Private Sub cmdDel_Click()
    On Error Resume Next
    DeleteSetting "Stop", txtUser
End Sub

Private Sub cmdExit_Click()
    If chkMode.Value Then
        SaveSetting "Stop", txtUser, "ScrollMode", InputBox("Введите ключ через знак `-`", "Key", "1 - 1")
    Else
          SaveSetting "Stop", txtUser, "ScrollMode", "False"
    End If
    Unload Form5
End Sub

Private Sub cmdSave_Click()
Rgs.Section = "Default"
Rgs.Key = "Start"
Rgs.Setting = chkStart.Value
Rgs.RegSave
If txtUser <> "" And txtPrl <> "" Then
    Rgs.Section = txtUser
    Rgs.Key = "Parol"
    Rgs.Setting = txtPrl
    Rgs.RegSave
    MsgBox "Ok. Parol is save", 0, "Information"
    
    Rgs.Section = txtUser
    Rgs.Key = "HelpParol"
    Rgs.Setting = txtHelp
    Rgs.RegSave
    MsgBox "Ok. Help is save", 0, "Information"
Else
    MsgBox "Введите имя и пароль", vbInformation, "Stop"
End If
Unload Form5
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Form1.Timer1.Enabled = False
Form3.tmrPrl.Enabled = False
Rgs.AppName = "Stop"
Rgs.Section = Form1.txtName
End Sub

