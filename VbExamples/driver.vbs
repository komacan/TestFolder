
Dim myArr(3)
myArr(0)=("C:\Users\Kurt\Documents\Unified Functional Testing\qtpObject1")
myArr(1)=("C:\Users\Kurt\Documents\Unified Functional Testing\qtpObject2")
myArr(2)=("C:\Users\Kurt\Documents\Unified Functional Testing\qtpObject3")
myArr(3)=("C:\Users\Kurt\Documents\Unified Functional Testing\qtpObject4")

Set objQTP= createobject("QuickTest.Application")
objQTP.Launch
'objQTP.Visible=true

For i = 0 To 3

objQTP.Open myArr(i)
objQTP.Test.Run
objQTP.Test.Close

Next
objQTP.Quit
Set objQTP= Nothing


