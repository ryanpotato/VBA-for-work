Attribute VB_Name = "ProperAndTrim"
Sub propercase()

Dim Jimmy As Range, cell As Range

Set Jimmy = Selection

For Each cell In Jimmy

cell.Value = WorksheetFunction.Proper(cell.Value)

Next cell

End Sub

Sub Trimmy()
Application.ScreenUpdating = False

Dim Billiam As Range
Set Billiam = Selection
For Each cell In Billiam
cell.Value = Trim(cell)
Next cell

Application.ScreenUpdating = True
End Sub
