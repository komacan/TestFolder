Set objXL=createobject("Excel.Application")
Set objWB= objXL.Workbooks.Open("C:\Users\Kurt\Desktop\VB\Book3.xlsx")
Set objWS= objWB.Worksheets("Sheet1")
'objXL.visible=True
objWS.Range("A1").Borders.LineStyle = 1 'Create a continuous border line
objWS.Range("A1").Borders.Color = RGB(255,0,0) 'Make the line black color
objWS.Range("A1").Borders.Weight = 4 'Make line thicker
objWS.cells(1,1).Font.Bold = TRUE 'Make Font Bold
objWS.cells(1,1).Font.Size = 24 'Make Font size 24
objWS.cells(1,1).Font.ColorIndex = 3 'Make Font color Red
objWB.Save
objWB.Close
objXL.Quit
Set objXL=Nothing
Set objWB1=Nothing
Set objWS1=Nothing