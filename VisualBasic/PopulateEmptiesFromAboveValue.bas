Attribute VB_Name = "PopulateEmptiesFromAboveValue"
Sub PopulateEmptyCellsWithValueDirectlyAbove()
Dim rng As Range
Set rng = Selection



For Each cell In rng


If cell = "" Then

'' copy and paste the value above
    cell.Offset(-1, 0).Select
    Selection.Copy
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
End If
Next cell



End Sub

