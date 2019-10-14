Attribute VB_Name = "DeleteSpacedRows"
Sub DeleteSpacedRows()
'
' Deletes rows spaced at selected interval
'

'
Dim delQuan As Integer
Dim movQuan As Integer
Dim i As Integer

MsgBox "This macro deletes rows from a table or workbook at a uniform interval. Make sure your table is uniform or you may lose data. Note that it sees hidden rows. Start by selecting a cell on a row you wanted deleted. Best to backup your data or do this on a copy as there's no 'undo' button for macros, hit cancel to cancel the operation."
delQuan = Application.InputBox("About how many rows to delete? Not a bad idea to underestimate and repeat the process. Hit cancel to abort")
movQuan = Application.InputBox("How many good rows in between the bad rows")



For i = 1 To delQuan
    ActiveCell.rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(movQuan, 0).Range("A1:E1").Select
Next i

End Sub


Sub DeleteSpacedRowswargs(delQuan As Integer, movQuan As Integer)
'
' Deletes rows spaced at selected interval
'

'
Dim delQuan As Integer
Dim movQuan As Integer
Dim i As Integer

MsgBox "This macro deletes rows from a table or workbook at a uniform interval. Make sure your table is uniform or you may lose data. Note that it sees hidden rows. Start by selecting a cell on a row you wanted deleted. Best to backup your data or do this on a copy as there's no 'undo' button for macros, hit cancel to cancel the operation."




For i = 1 To delQuan
    ActiveCell.rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(movQuan, 0).Range("A1:E1").Select
Next i

End Sub



