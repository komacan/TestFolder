'Mustafa Kurt

Set objQTP=createobject("QuickTest.Application") 'Create QTP object
Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
Set objTXT=objFSO.OpenTextFile("c:\VBSCRIPT\Text.txt",1) 'Open text file "TO READ"
objQTP.Visible=True ' Make the application visible during test
While objTXT.AtEndOfStream=False ' Do it while we have content
	objQTP.Open objTXT.ReadLine 'Open QTP file in the line
	objQTP.Test.Run ' Run the test inside QTP file
	obJQTP.Test.Close ' Close QTP
Wend 
objQTP.Quit 'Close QTP application
'Release Memory
Set objFSO=Nothing
Set objTXT=Nothing
