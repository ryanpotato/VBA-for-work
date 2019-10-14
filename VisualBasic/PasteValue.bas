Attribute VB_Name = "PasteValue"
 Sub PasteValue()

 
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False



 End Sub

Sub Today()
ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

End Sub
