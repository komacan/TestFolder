Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
objFSO.DeleteFile("c:\VBSCRIPT\text.txt") 'Delete text file"
objFSO.DeleteFolder("c:\VBSCRIPT") 'Delete the folder"
Set objFSO=Nothing