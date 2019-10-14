Attribute VB_Name = "Shortcuts"
Option Explicit

Sub To_FS()
'
' To_FS Macro
' This macro formats a cell to help with tie out to financial statements.
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormulaR1C1 = "To FS"
    With Selection.Font
       ' .Color = -16776961 'red
        .Color = -16727809 'orange
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
End Sub


Sub TB_link()
'
' TB_link Macro
' This macro fills a cell with tb link
'

'
    Selection.FormulaR1C1 = "TB link"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub arial8()
'
' arial8 Macro
'

'

    With Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
End Sub

Sub GoodNumberFormat()
'
' GoodNumberFormat Macro
' This macro creates an accounting number format with commas and parentheses but enables left or non-right justification in cells.
'

'
    Selection.NumberFormat = "_( #,##0_);_( (#,##0);_( ""-""??_);_(@_)"
    Selection.HorizontalAlignment = xlLeft
End Sub



Sub PBC()
'
' TB_link Macro
' This macro fills a cell with tb link
'

'
    Selection.FormulaR1C1 = "PBC"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
