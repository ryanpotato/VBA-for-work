Attribute VB_Name = "PivotTableReporterNS"
Option Explicit

Sub PivotTableReporterNS()
' Creates neat clean pivot table from Practice Management Reports
' This newer version avoids the .select method to optimize the code for faster performance
'' TODO Review for any bugs

Dim ClientName As String
Dim SaveSettingsFile As String
ClientName = Range("C10").FormulaR1C1
Dim BillPeriod As String

On Error GoTo ErrorHandler:
BillPeriod = Right(Range("A4").FormulaR1C1, Len(Range("A4").FormulaR1C1) - 5)
Dim DirToSave As String
'DirToSave = MyDocsPath & "\Documents\" & "EngagementTimeReports\"
Dim FileNameToSaveAs As String
FileNameToSaveAs = ClientName & BillPeriod
Application.EnableEvents = False
SaveSettingsFile = MyDocsPath() & "\Documents\" & "SaveSettings.txt"
Dim ReportWSName$
ReportWSName = Application.ActiveSheet.Name
Dim PTFormat As String
PTFormat = "PTFormat"

'make sure pivot table sheet doesn't already exist
If WorksheetExists("PivotTable") Then
MsgBox ("Pivot Table already exists. Delete Worksheet and re-run")
Exit Sub
End If

Application.ScreenUpdating = False

'format the data to be able to make the pivot table
FormatTable

'validate data (that the report didn't clip any hours off)
Call DataValidator(ReportWSName, PTFormat)

' make the pivot table on a second worksheet
MakePivot

' add the users custom cells to the pivot on that worksheet
AddUsercells

'Save the Workbook in selected directory
'TODO Get this working
'SaveWorkbook (ClientName)

Application.ScreenUpdating = True

' Create directory and save

    'check if chosen save directory exists, else make for persistence
    If Not FileExists(SaveSettingsFile) Then
    DirToSave = GetPermSaveFolder
    Call MakeSaveSettingsFile(DirToSave, SaveSettingsFile)
    Else
    DirToSave = ReadSaveLocation(SaveSettingsFile)
    End If

' save workbook automatically
Call SaveWorkbook(DirToSave, FileNameToSaveAs)

Application.EnableEvents = True

'wish them well
mymessage
Exit Sub
ErrorHandler:
MsgBox "Be sure the report worksheet is selected"
End Sub



Sub FormatTable()

Application.EnableEvents = False

' declare variables
Dim mycount As Integer
mycount = 0
Dim r As Long
Dim rangetosum As String
Dim valtocopy As String
Dim x2 As Range
Dim x3 As Range
Dim xdel As Range
Dim rowsdeleted As Integer
rowsdeleted = 0
Dim ReportWSName As String
ReportWSName = Application.ActiveSheet.Name
Dim PTFormat As String
PTFormat = "PTFormat"

'make sure pivot table sheet doesn't already exist
If Not WorksheetExists("PTFormat") Then
    Application.ActiveSheet.Copy After:=Application.ActiveSheet
    Application.ActiveSheet.Name = PTFormat
    Worksheets("PTFormat").Activate
Else:
    Worksheets("PTFormat").Activate

Exit Sub
End If

Dim GrandTotalHoursReport As Long
Dim GrandTotalHoursPT As Long


Dim i As Integer
Dim First As Range
Set First = Range("A2")
Dim LastCell As Range
Set LastCell = Application.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell)


' delete easy first Rs and Cs
Range("A1:A5").EntireRow.Delete
Range("B:D,G:H").Delete

' prepare loops
Set LastCell = LastCell.End(xlToLeft)
Set LastCell = LastCell.End(xlToLeft)
Set LastCell = LastCell.End(xlToLeft)
Set First = Range("A2")
mycount = Range(First, LastCell).Cells.count
First.Select
'GoTo tempstart

For i = 0 To mycount
' get names aligned
If ActiveCell.Offset(i).Font.Bold = False And InStr(ActiveCell.Offset(i).Value, "(") <> 0 And ActiveCell.Offset(i, 1) = "" Then
           
    'code for string to cut and paste names alongside data
    'ActiveCell.Offset(i).Cut
    
    ' deal with special case in Practice Management report formatting.
    'TODO make sure this is completely general - i.e. covers all cases
    If InStr(ActiveCell.Offset(i + 3, 0).Value, "Office ") <> 0 Then
        
        valtocopy = ActiveCell.Offset(i)
        Set x3 = ActiveCell.Offset(i + 1)
        x3.Formula = valtocopy
        x3.Font.FontStyle = "arial"
        x3.Font.Size = 8
        '.ActiveCell.Offset(i).Copy (ActiveCell.Offset(i + 1))
    Else
    valtocopy = ActiveCell.Offset(i).FormulaR1C1
        'ActiveCell.Offset(i + 2).Value = valtocopy
        'ActiveCell.Offset(i).Copy (Range("a5"))

    
    Set x2 = ActiveCell.Offset(i + 2)
    End If
Repeater:
        'ActiveSheet.Offset(i + 2).Paste
              
        
        x2.Formula = valtocopy
        x2.Font.FontStyle = "arial"
        x2.Font.Size = 8
        If x2.Offset(1, 1) <> "" Then
            Set x2 = x2.Offset(1, 0)
            x2.Formula = valtocopy
            GoTo Repeater
    
        'Else
        ' code to move to next
           ' x2 = x2.Offset(1, 0)
        End If
Else
'do nothing
End If

Next i

'go back to startcell
First.Offset(-1).Select

'loop to delete unwanted rows (empties and most bolds)
tempstart:


For i = 1 To mycount
    Set xdel = ActiveCell.Offset(i - rowsdeleted)
    If xdel.Row > LastCell.Row Then
        GoTo LastIteration
    End If
    
    'Now delete all unwanted rows (blanks and most bold rows)
    If (xdel.Offset(0, 1).FormulaR1C1 = "" Or xdel.Font.Bold = True) And (InStr(xdel.FormulaR1C1, "Totals") = 0) Then
'Again:
        xdel.EntireRow.Delete
        rowsdeleted = rowsdeleted + 1
   '     If (ActiveCell = "" And ActiveCell.Offset(0, 1) = "") Or (ActiveCell.Font.Bold = True And InStr(ActiveCell.Value, "Totals") = 0) Then
           ' GoTo Again
     '   End If

    ElseIf (xdel.Font.Bold = True And InStr(xdel.FormulaR1C1, "Totals") <> 0) Then
        rangetosum = xdel.Offset(-1, 2).Address & ":" & xdel.Offset(-1, 2).End(xlUp).Address
        xdel.Offset(0, 2).Formula = "=SUM(" & rangetosum & ")"
    '' Recalc total bill hrs for data validation purposes
    
    End If
    'ActiveCell.Offset(1).Select
LastIteration:
Next i
Range("A1").Select

End Sub

Sub MakePivot()

'' Code below from source https://excelchamps.com/blog/vba-to-create-pivot-table/
'' Copyright by them as template if they have any, else just credit where due

'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long

'create new Worksheet
On Error Resume Next
Set DSheet = Application.ActiveSheet
Application.DisplayAlerts = False
Sheets.Add After:=Application.ActiveSheet
Application.ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
'Set DSheet = Worksheets("Report")

'Define Data Range
LastRow = DSheet.Cells(rows.count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow - 1, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(5, 1), _
TableName:="TimeUsage")


'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="TimeUsage")


'Insert Row Fields
With Application.ActiveSheet.PivotTables("TimeUsage").PivotFields("Service Description")
    .Orientation = xlRowField
    .Position = 1
End With

'Insert Column Fields
With Application.ActiveSheet.PivotTables("TimeUsage").PivotFields("Employee Name (Number)")
    .Orientation = xlColumnField
    .Position = 1
End With

'Insert Data Field
With Application.ActiveSheet.PivotTables("TimeUsage").PivotFields("Bill Hrs")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    ''' TODO try toggling this to see what happens
    '.NumberFormat = "#,##0"
    .Name = "Sum of Bill Hours"
End With

'Format Pivot
'' TODO Figure out what it does and if is necessary
'TableActiveSheet.PivotTables("TimeUsage").ShowTableStyleRowStripes _
'= TrueActiveSheet.PivotTables("TimeUsage").TableStyle2 = "PivotStyleMedium9"
Application.ActiveSheet.Cells.ColumnWidth = 9.8

End Sub

Sub AddUsercells()
Dim FRowNum As Integer
Dim LRowNum As Integer
Dim MyArray(3) As String
Dim TimesToDo As Integer
MyArray(0) = "Budget"
MyArray(1) = "Prior Year"
MyArray(2) = "Budget to Actual"
MyArray(3) = "PY to CY"
Dim i As Integer

Application.ActiveSheet.Range("A1").Select
Selection.End(xlDown).End(xlDown).Select
LRowNum = ActiveCell.Row
Selection.End(xlUp).Select
FRowNum = ActiveCell.Row
TimesToDo = LRowNum - FRowNum - 2
ActiveCell.Offset(1, 1).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
For i = 0 To 3
    Selection.FormulaR1C1 = MyArray(i) ' "Budget"
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.Offset(0, 1).Select
Next i

'shade yeller
ActiveCell.Offset(1, -4).Select
ActiveCell.Range("A1:B" & CStr(TimesToDo)).Select
''TODO make conditional yellow
BlankYellow
'Selection.Interior.Color = 65535
ActiveCell.Offset(0, 2).Select

For i = 0 To TimesToDo
     ActiveCell.Formula = "=" & ActiveCell.Offset(0, -2).Address & "-" & ActiveCell.Offset(0, -3).Address
     Selection.Style = "Comma"
     Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
     '' todo make conditional red, green, orange
     FormatCellsROG
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Formula = "=" & ActiveCell.Offset(0, -2).Address & "-" & ActiveCell.Offset(0, -4).Address
     Selection.Style = "Comma"
     Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    '' todo make conditional red, green, orange
    FormatCellsROG
     ActiveCell.Offset(1, -1).Range("A1").Select
Next i
Selection.Offset(-1, -2).Select
For i = 0 To 1
    ActiveCell.Formula = "=SUM(R[-" & CStr(TimesToDo) & "]C:R[-1]C)"
    ActiveCell.Offset(0, 1).Range("A1").Select
Next i

End Sub

Sub mymessage()
MsgBox ("Have a nice day")

End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Sub SaveWorkbook(DirToSave As String, FileName As String)
Dim PathToSave As String
FileName = Replace(FileName, "/", "-")
PathToSave = DirToSave & "\" & FileName

'    If Len(Dir(PathToSave)) <> 0 Then
'    MsgBox ("File already exists. Overwrite?")
'    End If
 On Error Resume Next
ActiveWorkbook.SaveAs (PathToSave)
End Sub


Public Function MyDocsPath() As String
MyDocsPath = VBA.Environ$("USERPROFILE")
End Function

Sub FormatCellsROG()

'' Adds conditional formatting for cells based on value differentials

    '' sets red shade RED text if less than -5
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=-5"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    '' sets ORANGE shade and red text if between -5 and -.51
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-5", Formula2:="=-.51"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    '' sets GREEN shade and black text if greater than 0
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

Sub BlankYellow()
Dim cell As Range
    For Each cell In Selection
    cell.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(" & cell.Address & "))=0"
    cell.FormatConditions(cell.FormatConditions.count).SetFirstPriority
    With cell.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
    End With
    cell.FormatConditions(1).StopIfTrue = False
    Next cell

End Sub

Sub DirectoryExistenceCheck(strFolderName As String)

Dim strFolderExists As String

    strFolderExists = Dir(strFolderName, vbDirectory)
 
    If Len(Dir(strFolderName, vbDirectory)) = 0 Then
    MkDir strFolderName
    End If

End Sub

Sub DataValidator(ReportName As String, PTData As String)

Dim rng1 As Range
Dim rng2 As Range
Dim GT1 As Single
Dim GT2 As Single
Set rng1 = Application.ActiveSheet.UsedRange
Dim rng3 As Range

Set rng3 = rng1.Find("Grand Totals").Offset(0, 2)
GT1 = rng3.Value
Set rng2 = Application.Worksheets(ReportName).UsedRange
GT2 = rng2.Find("Grand Totals").Offset(0, 5).Value

If GT1 - GT2 <> 0 Then
    rng3.Select
    rng3.Interior.ColorIndex = 8
    MsgBox ("Warning, values don't match report")
    End
End If

Range("A1").Select
End Sub

Sub MakeSaveSettingsFile(DirToRecord As String, FileToRecordTo As String)

        Open FileToRecordTo For Output As #1
        Write #1, DirToRecord
        Close #1
        
End Sub
Function GetPermSaveFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Dim strpath As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    '.InitialFileName = strpath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetPermSaveFolder = sItem
'GetPermSaveFolder = fldr
End Function

Function ReadSaveLocation(SaveSettingsFile As String) As String
Dim content As String

        Open SaveSettingsFile For Input As #1
        Input #1, content
        Close #1
ReadSaveLocation = content
End Function

Function FileExists(FilePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function





