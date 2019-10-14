Attribute VB_Name = "GCR"
Sub gcr()
   ActiveCell.FormulaR1C1 = "GCR"
End Sub

Sub UserNameFullPaste()
   ActiveCell.FormulaR1C1 = Application.username
End Sub

Sub UserNameInitialsPaste()
   ActiveCell.FormulaR1C1 = UCase((Environ$("Username")))

End Sub
