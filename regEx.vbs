Function sendEmail() 
email = InputBox("Enter email: ")

Set re = New RegExp

With re
.Pattern = "^[\w]{1,}[[\.]{0,}[\w]{1,}]{0,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
.IgnoreCase = False
.Global = False
End With

If re.Test(email) Then
MsgBox("Crash Report Sent")
Else
MsgBox("Please enter a valid email address")
Call sendEmail()
End If
End Function

a = MsgBox("Crashed!!! Do you want to send a crash report?", 4, "Crashed")

If a = 6 Then
Call sendEmail()

ElseIf a = 7 Then
MsgBox("Crash Report not Sent")
End If



