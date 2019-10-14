Attribute VB_Name = "RegularNames"
Sub RegularNames()
Dim WkRg As Range, Rg As Range
Dim T
    Set WkRg = Selection
    For Each Rg In WkRg
        T = Split(Rg, ",")
        Rg = Trim(T(1)) & " " & Trim(T(0))
    Next Rg
End Sub

'' Note the following two macros didn't work but they're there for legacy and to maybe someday debug


Sub NameFormat()
'
' NameFormat Macro
' Changes name from last, first to first last (e.g. "Smith, John" to "John Smith")
'

'
Dim MyFormulaString As String
Dim rngMyRange As Range
Set rngMyRange = Selection
OffsetRange = rngMyRange.Offset(0, -1)


MyFormulaString = "=RIGHT(A5,LEN(A5)-FIND("","",A5))&"" ""&LEFT(A5,FIND("","",A5)-1)"
'MyFormulaString = "=RIGHT(OffsetRange,LEN(OffsetRange)-FIND("","",OffsetRange))&"" ""&LEFT(OffsetRange,FIND("","",OffsetRange)-1)"
rngMyRange.Formula = MyFormulaString

End Sub
Sub NameFormat2()
Dim MyFormulaString As String
Dim rngMyRange As Range

For Each rngMyRange In [A1:A100]  'raw data in col A
    If rngMyRange.Value <> "" Then
        MyFormulaString = "=trim(RIGHT(" & rngMyRange.Address & ",LEN(" & rngMyRange.Address & ")-FIND("",""," & rngMyRange.Address & "))&"" ""&LEFT(" & rngMyRange.Address & ",FIND("",""," & rngMyRange.Address & ")-1))"
        rngMyRange.Offset(0, 1).Formula = MyFormulaString 'results goes to col B (A offset + 1)
    End If
Next rngMyRange
End Sub
Option Explicit

Option Explicit



