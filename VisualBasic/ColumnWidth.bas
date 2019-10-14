Attribute VB_Name = "ColumnWidth"
Sub cw()
Dim cw As Double
Dim neww As Double

cw = ActiveCell.ColumnWidth()
If MsgBox("The column width is: " & cw & ". Would you like to change it?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
neww = Application.InputBox("input your new desired column width")
ActiveCell.ColumnWidth = neww


End Sub

