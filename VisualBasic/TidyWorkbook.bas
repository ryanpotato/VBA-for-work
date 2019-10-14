Attribute VB_Name = "TidyWorkbook"

Sub UnFreeze()
'Update 20140317
Dim Ws As Worksheet
Application.ScreenUpdating = False
For Each Ws In Application.ActiveWorkbook.Worksheets
    Ws.Activate
    With Application.ActiveWindow
            .FreezePanes = False
            .Zoom = 100
            .DisplayGridlines = False
    
    Ws.Cells.Font.Size = "8"
    Ws.[a1].Select
    ActiveWorkbook.Worksheets(1).Activate
    End With
Next
Application.ScreenUpdating = True
End Sub
