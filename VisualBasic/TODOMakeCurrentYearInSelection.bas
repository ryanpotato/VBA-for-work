Attribute VB_Name = "TODOMakeCurrentYearInSelection"
Sub MakeCurrentYearInSelection()

''FANCIER PSEUDO

''copy contents of selection as string
''CellContents = Selection.copy
''CurrentYear as str; set CurrentYear = GetCurrentYear()
''search inside selection for current year minus one
    ''count number of times found and create a For Loop
        ''YearInQuestionInstances = [how to count]
        ''For i=1, YearInQuestionInstances, i++
            ''Get Instring Position
            ''InStr( [start], string, substring, [compare] )
            ''YPosition = InStr(1 + YPosition, CellContents, CurrentYear )
                ''if YPosition !=0
                    ''replace prior year with current year, keeping rest intact
                    
        ''Next i
''copy-paste as value/overwrite cell

'' alternatively search for any year and increment one, no matter what. E.g. py becomes cy, cy becomes fy+1



''SIMPLER PSEUDO
''copy contents of selection as string
''CurrentYear as str; set CurrentYear = GetCurrentYear()
''search inside selection for current year minus one
''if found, get location in string
    ''if found, replace prior year with current year, keeping rest intact
    ''paste as value
    

'' alternatively search for any year and increment one, no matter what. E.g. py becomes cy, cy becomes fy+1
End Sub

