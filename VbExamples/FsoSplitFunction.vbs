Set objFSO=createobject("Scripting.FileSystemObject") 'Create the FSO object
'Set objTS = objFSO.OpenTextFile("C:\Study\TestCase.txt",1) 'Read from the TESTCASE file
Set objTD = objFSO.OpenTextFile("C:\Study\TestData.txt",1) 'Open to READ
Set objOT = objFSO.OpenTextFile("C:\Study\Output.txt",2,true) 'Open to write, checks if the file exist
While objTD.AtEndOfStream=False '
	str1=split(objTD.Readline,":")
	str =split(str1(1),",")
	objOT.WriteLine "First Name="&str(0)&", Last Name="&str(1)
Wend
Set objFSO=Nothing
Set objTD=Nothing
Set objOT=Nothing
