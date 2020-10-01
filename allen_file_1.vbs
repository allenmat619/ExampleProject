
Function file()
Set obj = createobject("Scripting.FileSystemObject")

a = InputBox("What do you want to do (Enter number)?		1 : Create File, 2 : Delete File, 3 : Write to File, 4 : Read from File")

If a = 1 Then
cName = InputBox("Enter name of the file you want to create")
cFile = "C:\Users\u25u85\Desktop\VBScript\" & cName & ".txt"
obj.CreateTextFile cFile
Set obj = Nothing
MsgBox("File created successfully")

ElseIf a = 2 Then
dName = InputBox("Enter the name of the file you want to delete")
dFile = "C:\Users\u25u85\Desktop\VBScript\" & dName & ".txt"                                         
obj.DeleteFile dFile
Set obj = Nothing        
MsgBox("File deleted successfully")

ElseIf a = 3 Then
wName = InputBox("Enter the name of the file you want to write to")
Set objFile = obj.OpenTextFile("C:\Users\u25u85\Desktop\VBScript\" & wName & ".txt", 2)
wLine = InputBox("Enter the line you want to write into the file")
objfile.WriteLine(wLine)
objFile.Close
Set obj = Nothing

ElseIf a = 4 Then
rName = InputBox("Enter the name of the file which you want to read from")
Set objFile = obj.OpenTextFile("C:\Users\u25u85\Desktop\VBScript\" & rName & ".txt", 1)
str = objFile.ReadAll
MsgBox str

Else
MsgBox("Wrong choice. Please try again")
End If
cont = MsgBox("Do you want to continue?", 4)
If cont = 6 Then
Call file()
End If
End Function


Call file()



