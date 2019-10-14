Attribute VB_Name = "InsertRow"
Sub insertrow()
'
' insertrow Macro
'

Dim MySelection As Range
Dim RowToAdd As Range

    Set MySelection = Selection
    Set RowToAdd = MySelection.rows.EntireRow
    Selection.Insert Shift:=xlDown

End Sub
