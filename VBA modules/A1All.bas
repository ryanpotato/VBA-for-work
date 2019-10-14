Attribute VB_Name = "A1All"
Sub A1AllWorksheets()
'Update 20140317
Dim Ws As Worksheet
'Application.ScreenUpdating = False
For Each Ws In Application.ActiveWorkbook.Worksheets
    Ws.Activate
    With Application.ActiveWindow
    Ws.[a1].Select
    End With
Next
ActiveWorkbook.Worksheets(1).Activate
'Application.ScreenUpdating = True
End Sub
