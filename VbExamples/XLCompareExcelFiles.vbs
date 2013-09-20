Set objXL=createobject("Excel.Application")
Set objWB1= objXL.Workbooks.Open("C:\Users\Kurt\Desktop\VB\Book1.xlsx")
Set objWB2= objXL.Workbooks.Open("C:\Users\Kurt\Desktop\VB\Book2.xlsx")
Set objWS1= objWB1.Worksheets("Sheet1")
Set objWS2= objWB2.Worksheets("Sheet1")

vT1RowCount = objWS1.UsedRange.Rows.count 'Number of Rows in Book1-Sheet1
vT1ColumnCount = objWS1.UsedRange.Columns.count 'Number of Columns in Book1-Sheet1
vT2RowCount = objWS2.UsedRange.Rows.count 'Number of Rows in Book2-Sheet1
vT2ColumnCount = objWS2.UsedRange.Columns.count  'Number of Columns in Book2-Sheet1
vCounter=0

If vT1RowCount=vT2RowCount and vT1ColumnCount=vT2ColumnCount Then 'Check if Column and Row counts are matching
    For i = 1 To vT1RowCount
        For j = 1 To vT1ColumnCount
            If objWS1.Cells(i,j)=objWS2.Cells(i,j) Then 'Compare each cell on both Excel files
            Else
                vCounter=vCounter+1 'If any cell doesnt match Add 1 to vCounter
            End If
        Next
    Next
    If vCounter=0 Then
        Msgbox "Excel files are same"
    Else
        Msgbox "Excel files are different"
    End If
Else
    Msgbox "Excel files are different"
End If

objWB1.Close
objWB2.Close
objXL.Quit
Set objXL=Nothing
Set objWB1= Nothing
Set objWB2= Nothing
Set objWS1= Nothing
Set objWS2= Nothing
