Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
If objFSO.FolderExists("c:\VBSCRIPT")=False Then
	objFSO.CreateFolder("c:\VBSCRIPT") 'Create folder
End If
objFSO.CreateTextFile("c:\VBSCRIPT\text.txt") 'Create the text file
objFSO.CreateTextFile("c:\VBSCRIPT\text1.txt") 'Create the text file
objFSO.CreateTextFile("c:\VBSCRIPT\text2.txt") 'Create the text file
Set objTXT=objFSO.OpenTextFile("c:\VBSCRIPT\text.txt",2,true) 'Open the text file "TO WRITE" if the file doesnt exist, CREATE the file
Set objTXT1=objFSO.OpenTextFile("c:\VBSCRIPT\text1.txt",2,true)
Set objTXT2=objFSO.OpenTextFile("c:\VBSCRIPT\text2.txt",2,true)
objTXT.WriteLine "This is me writting to a text I created a second ago" 'Write to the text

For i = 1 To 20 ' "i" number of Lines will be written
	objTXT1.WriteLine "This is Text1, Line "&i  'Write to the text
Next

objTXT2.WriteLine "This is Text2 Line 1" 'Write to the text
objTXT2.WriteLine "This is Text2 Line 2" 'Write to the text

Set objFSO=Nothing
Set objTXT=Nothing
Set objTXT1=Nothing
Set objTXT2=Nothing