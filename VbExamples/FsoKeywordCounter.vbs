Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
Set ObjTXT=objFSO.OpenTextFile("c:\VBSCRIPT\Text1.txt",1) 'Open text file "TO READ"
vKeyword="2"
vCounter=0
While objTXT.AtEndOfStream=False
myArr= split(objTXT.ReadLine)
For i = 0 To Ubound(myArr)
	If myArr(i)=vKeyword Then
		vCounter=vCounter+1
	End If
Next
Wend
If vCounter = 0 Then
	MsgBox "Keyword='"&vKeyword&"' does not exist "
Else
	MsgBox vCounter&" Keyword='"&vKeyword&"' is found in the text file"
End If