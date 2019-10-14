Attribute VB_Name = "DeleteHiddenRows"
Sub DeleteHiddenRows()
Dim SelectedArea As Range
Set SelectedArea = Selection

For Each rw In SelectedArea.rows

If rows(rw.Row).Hidden = True Then rows(rw.Row).EntireRow.Delete
Next

End Sub
