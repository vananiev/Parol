Attribute VB_Name = "Ferst"
Option Explicit
Dim A As New Scripting.FileSystemObject

Sub Main()
    
    Open App.Path & "\starting.reg" For Output As #1
        Print #1, "REGEDIT4" & vbCrLf
        Print #1, "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run]"
        Print #1, """Stop""" & "=" & """" & Replace(A.GetSpecialFolder(SystemFolder) & "\parol.exe", "\", "\\") & """" & vbCrLf
        Print #1, "[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Stop\Default]"
        Print #1, """Start""" & "=" & """0""" & vbCrLf
   Close #1
End Sub
