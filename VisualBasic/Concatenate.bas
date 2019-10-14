Attribute VB_Name = "Concatenate"
Option Explicit

'The following 4 macros are used to call the Concatenate_Formula macro.
'The Concatenate_Formula macro has different options, and these 4 macros
'run the Concatenate_Formula macro with different options.  You will want
'to assign any of these macros to a ribbon button or keyboard shortcut.

Sub Ampersander()
    'Creates a basic Ampersand formula with no options
    Call Concatenate_Formula(False, False)
End Sub

Sub Ampersander_Options()
    'Creates an Ampersand formula and prompts the user for options
    'Options are absolute refs and separator character
    Call Concatenate_Formula(False, True)
End Sub

Sub Concatenate()
    'Creates a basic CONCATENATE formula with no options
    Call Concatenate_Formula(True, False)
End Sub

Sub Concatenate_Options()
    'Creates an CONCATENATE formula and prompts the user for options
    'Options are absolute refs and separator character
    Call Concatenate_Formula(True, True)
End Sub


Sub Concatenate_Formula(bConcat As Boolean, bOptions As Boolean)

Dim rselected As Range
Dim c As Range
Dim sArgs As String
Dim bCol As Boolean
Dim bRow As Boolean
Dim sArgSep As String
Dim sSeparator As String
Dim rOutput As Range
Dim vbAnswer As VbMsgBoxResult
Dim lTrim As Long
Dim sTitle As String

    'Set variables
    Set rOutput = ActiveCell
    bCol = False
    bRow = False
    sSeparator = ""
    sTitle = IIf(bConcat, "CONCATENATE", "Ampersand")
    
    'Prompt user to select cells for formula
    On Error Resume Next
    Set rselected = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
    On Error GoTo 0
    
    'Only run if cells were selected and cancel button was not pressed
    If Not rselected Is Nothing Then
        
        'Set argument separator for concatenate or ampersand formula
        sArgSep = IIf(bConcat, ",", "&")
        
        'Prompt user for absolute ref and separator options
        If bOptions Then
        
            vbAnswer = MsgBox("Columns Absolute? $A1", vbYesNo)
            bCol = IIf(vbAnswer = vbYes, True, False)
            
            vbAnswer = MsgBox("Rows Absolute? A$1", vbYesNo)
            bRow = IIf(vbAnswer = vbYes, True, False)
                
            sSeparator = Application.InputBox(Prompt:= _
                        "Type separator, leave blank if none.", _
                        Title:=sTitle & " separator", Type:=2)
        
        End If
        
        'Create string of cell references
        For Each c In rselected.Cells
            sArgs = sArgs & c.Address(bRow, bCol) & sArgSep
            If sSeparator <> "" Then
                sArgs = sArgs & Chr(34) & sSeparator & Chr(34) & sArgSep
            End If
        Next
        
        'Trim extra argument separator and separator characters
        lTrim = IIf(sSeparator <> "", 4 + Len(sSeparator), 1)
        sArgs = Left(sArgs, Len(sArgs) - lTrim)

        'Create formula
        'Warning - you cannot undo this input
        'If undo is needed you could copy the formula string
        'to the clipboard, then paste into the activecell using Ctrl+V
        If bConcat Then
            rOutput.Formula = "=CONCATENATE(" & sArgs & ")"
        Else
            rOutput.Formula = "=" & sArgs
        End If
        
    End If

End Sub

