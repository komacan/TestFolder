Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
Set ObjTXT=objFSO.OpenTextFile("c:\VBSCRIPT\text2.txt",1) 'Open text file "TO READ"
While objTxt.AtEndOfStream=False ' Do it while we have content
msgbox objTXT.ReadLine
Wend
Set objFSO=Nothing
Set objTXT=Nothing