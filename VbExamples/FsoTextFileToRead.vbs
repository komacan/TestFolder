Set objFSO= createobject("Scripting.Filesystemobject")
Set objTXT = objFSO.OpenTextFile("C:\VBSCRIPT\text1.txt",1) 'Open Text file "TO READ"
Msgbox objTXT.ReadLine
Set objFSO=nothing
Set objTXT=nothing

