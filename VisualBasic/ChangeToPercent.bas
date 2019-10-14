Attribute VB_Name = "ChangeToPercent"
Sub changetopercent()
'
' changetopercent Macro

Dim Billy As Range, john As Range
Set Billy = Selection

For Each john In Billy
If Not IsEmpty(john) Then
    john = john / 100
    john.Style = "Percent"  'Style is a built-in function
End If
Next
End Sub

























Sub change2percent2()
Dim n1 As Range
Dim n2 As Range

    Set n1 = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
        Set n2 = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
                    
                    
                    
End Sub
