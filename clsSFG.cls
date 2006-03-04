VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get SFG() As Variant
    SFG = vSafeguard
End Property

Public Property Let SFG(ByVal vNewValue As Variant)
    vSafeguard = vNewValue
End Property


Public Property Let Timer(ByVal vNewValue As Boolean)
    frmSFG.tmrWait.Enabled = vNewValue
End Property
