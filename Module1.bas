Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    On Error Resume Next
    
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description
    Else
        Form1.Show
    End If
        
End Sub
