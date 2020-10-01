
Do
a = MsgBox("Sorry an Error Occurred",5,"Error Occurred")
If a = 2 Then
a = MsgBox("Do you want to submit an Error Report?", 4, "Submit Error")
If a = 6 Then
a = MsgBox("Error Report was submitted", 0, "Submitted")
ElseIf a = 7 Then
a = MsgBox("Error Report was not submitted", 0, "Not Submitted")
End If
End If
Loop While a = 4