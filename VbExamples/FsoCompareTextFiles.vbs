Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
Set objTXT1=objFSO.OpenTextFile("c:\VBSCRIPT\Text1.txt",1) 'Open text file "TO READ"
Set objTXT2=objFSO.OpenTextFile("c:\VBSCRIPT\Text2.txt",1) 'Open text file "TO READ"
If objTXT2.ReadAll=objTXT1.ReadAll Then
	Msgbox "Files are identical"
	Else
	Msgbox "Files are different"
End If
Set objFSO=Nothing
Set objTXT1=Nothing
Set objTXT2=Nothing
