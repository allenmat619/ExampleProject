

a = MsgBox ("Do you want to run a script?", 4)

If a = 6 Then

Dim obj

Set obj = Wscript.CreateObject("WScript.Shell")

obj.Run "allen_file_1.vbs"

Set obj = Nothing

Else
MsgBox("Script was not run")

End If