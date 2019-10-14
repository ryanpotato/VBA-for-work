Attribute VB_Name = "arrows"
Option Explicit


Sub drawarrows(FromRange As Range, ToRange As Range, Optional RGBcolor As Long, Optional LineType As String)
'---------------------------------------------------------------------------------------------------
'---Script: DrawArrows------------------------------------------------------------------------------
'---Created by: Ryan Wells -------------------------------------------------------------------------
'---Date: 10/2015-----------------------------------------------------------------------------------
'---Description: This macro draws arrows or lines from the middle of one cell to the middle --------
'----------------of another. Custom endpoints and shape colors are suppported ----------------------
'---------------------------------------------------------------------------------------------------
    
Dim dleft1 As Double, dleft2 As Double
Dim dtop1 As Double, dtop2 As Double
Dim dheight1 As Double, dheight2 As Double
Dim dwidth1 As Double, dwidth2 As Double
dleft1 = FromRange.Left
dleft2 = ToRange.Left
dtop1 = FromRange.Top
dtop2 = ToRange.Top
dheight1 = FromRange.Height
dheight2 = ToRange.Height
dwidth1 = FromRange.Width
dwidth2 = ToRange.Width
 
ActiveSheet.Shapes.AddConnector(msoConnectorStraight, dleft1 + dwidth1 / 2, dtop1 + dheight1 / 2, dleft2 + dwidth2 / 2, dtop2 + dheight2 / 2).Select
'format line
With Selection.ShapeRange.Line
    .BeginArrowheadStyle = msoArrowheadNone
    .EndArrowheadStyle = msoArrowheadOpen
    .Weight = 1.5
    .Transparency = 0.5
    If UCase(LineType) = "DOUBLE" Then 'double arrows
        .BeginArrowheadStyle = msoArrowheadOpen
    ElseIf UCase(LineType) = "LINE" Then 'Line (no arows)
    .EndArrowheadStyle = msoArrowheadNone

    ElseIf UCase(LineType) = "SINGLE" Then
    .EndArrowheadStyle = msoArrowheadTriangle
            Else 'single arrow
        'defaults to an arrow with one head
    End If
    'color arrow
    If RGBcolor <> 0 Then
        .ForeColor.RGB = RGBcolor 'custom color
    Else
        .ForeColor.RGB = RGB(255, 0, 0)     'red (DEFAULT)
    End If
End With
 
End Sub

Sub DrawAcrossSelection()
Call drawarrows(Selection.Cells(1, 1), Selection.Cells(Selection.rows.count, Selection.Columns.count))
End Sub


'Sub DrawAcrossSelectionblue()
'Call drawarrows(Selection.Cells(1, 1), Selection.Cells(Selection.Rows.Count, Selection.Columns.Count),)
'End Sub



Sub DrawBlueAcrossSelection()
Dim Blu As Long
Blu = RGB(0, 0, 255)
Call drawarrows(Selection.Cells(1, 1), Selection.Cells(Selection.rows.count, Selection.Columns.count), Blu)
End Sub


Sub Fence()
'
' Fence Macro
' Creates fence
'
' Keyboard Shortcut: Ctrl+Shift+F
'

    ActiveCell.SpecialCells(xlLastCell).Select
    ActiveCell.Offset(2, 2).Range("A1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16711680
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.SpecialCells(xlLastCell).Select
    Range(Selection, Selection.End(xlUp)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 16711680
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 2
    Range("A1").Select
End Sub


Sub BoldRed()
'
' BoldRed Macro
' Creates bold red size 9 font for tickmarking
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    With Selection.Font
        .Bold = True
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub
Sub Ariall()
'
' Arial9 Macro
' Makes all cells arial 9
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    Cells.Select
    With Selection.Font
        .Name = "Arial"
    End With
End Sub

Sub Nine()

    Cells.Select
    With Selection.Font
    .Size = 9
    End With

End Sub

Sub boldblue()

Selection.Font.Bold = True
Selection.Font.Color = -65536


End Sub

Sub MakeBigger()
    Selection.Font.Size = Selection.Font.Size + 1
End Sub
Sub MakeSmaller()
    Selection.Font.Size = Selection.Font.Size - 1
End Sub

' Delete Empty Rows


Sub foo()
  Dim r As Range, rows As Long, i As Long
  Set r = ActiveSheet.Range("A1:Z50")
  rows = r.rows.count
  For i = rows To 1 Step (-1)
    If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
  Next
End Sub

Sub disablePageBreaks()
ActiveSheet.DisplayPageBreaks = False
End Sub
Sub ReverseSign()
'
' ReverseSign Macro
' Reverses the sign of a cell when run
'

Dim c As Range
For Each c In Selection
c.Value = -c.Value
Next c
End Sub
   



Sub lockCellsWithFormulas()
With ActiveSheet
.Unprotect
.Cells.Locked = False
.Cells.SpecialCells(xlCellTypeFormulas).Locked = True
.Protect AllowDeletingRows:=True

End With
End Sub
'Sub GoalSeekVBA()
Dim Target As Long
On Error GoTo ErrorHandler
Target = InputBox("Enter the required value", "Enter Value")
'Worksheets("Goal_Seek").Activate

With ActiveWorkbook.ActiveSheet
.Range(Selected.Cells).GoalSeek _
Goal:=Target, _
ChangingCell:=Range("C2")
End With
Exit Sub
ErrorHandler:
MsgBox ("Sorry, value is not valid.")
End Sub
Sub highlightMaxValue()
Dim rng As Range
For Each rng In Selection
If rng = WorksheetFunction.Max(Selection) Then
rng.Style = "Good"
End If
Next rng
End Sub

Sub highlightMinValue()



Dim rng As Range
    For Each rng In Selection
        If rng = WorksheetFunction.Min(Selection) Then
        rng.Style = "Good"
        End If
    Next rng

End Sub



