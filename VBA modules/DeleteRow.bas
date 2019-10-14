Attribute VB_Name = "DeleteRow"
Sub DeleteRow()
'
' DeleteRow Macro
'

'
Dim MySelection As Range
Dim RowsToDelete As Range

    Set MySelection = Selection
    Set RowsToDelete = MySelection.rows.EntireRow
    'Set RowToDelete = ActiveCell.rows("1:1").EntireRow
    RowsToDelete.Delete
    
End Sub
 

