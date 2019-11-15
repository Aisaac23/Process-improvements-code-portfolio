Sub Compliance_fill() 'Last update 09/29/2016

'*****Variables definition******
    Dim range_fill, range_get As String
    Dim count, rows_to_fill As Integer
    Dim SourceRange, fillRange As Range
    'We insert the formula in the cell D6
    Worksheets("Compliance").Range("D6").Formula = "=IFNA(IF(ISNUMBER(MATCH($C6,INDIRECT(""'[Audits.xlsx]Sheet1'!"" & TEXT([Audits.xlsx]Sheet1!$A$51,"""")), 0)), INDIRECT(""'[Audits.xlsx]Sheet1'!"" & SUBSTITUTE(ADDRESS(1, MATCH(""*""&D$4&""*"", [Audits.xlsx]Sheet1!$A$1:$DD$1,0)+1,4 ),""1"", """") & MATCH($C6,INDIRECT(""'[Audits.xlsx]Sheet1'!"" & TEXT([Audits.xlsx]Sheet1!$A$51,"""")), 0)+1), """"), """")"
    ' Filling vertically
    rows_to_fill = Range("A6").End(xlDown).Row
    range_fill = "D6" & ":D" & CStr(rows_to_fill)
    Set SourceRange = Worksheets("Compliance").Range("D6")
    Set fillRange = Worksheets("Compliance").Range(range_fill)
    SourceRange.AutoFill Destination:=fillRange
    'By rows
    count = 6
    While count <= Range("A6").End(xlDown).Row
        range_get = "D" & CStr(count)
        range_fill = "D" & CStr(count) & ":BZ" & CStr(count)
        Set SourceRange = Worksheets("Compliance").Range(range_get)
        Set fillRange = Worksheets("Compliance").Range(range_fill)
        SourceRange.AutoFill Destination:=fillRange 'the := colon equal sign symbols. This tells VBA that we are setting the parameter to a value or object. This is useful for seting optional parameters. 
        count = count + 1
    Wend
    '***************************
    'For the Non-Prod environments
    '***************************
    Dim range_rows As String
    count = Range("A6").End(xlDown).Row + 2
    range_rows = "A" & CStr(count)
    range_get = "D" & CStr(count)
	'We obtain the range and insert the corresponding formula into it
    Worksheets("Compliance").Range(range_get).Formula = "=IFERROR(IF(ISNUMBER(MATCH($A27, $A$6:$A$25,0)), INDIRECT(TEXT(ADDRESS(MATCH($A27, $A$6:$A$25,0)+5, COLUMN(D$4),1), """")), """"), """")"
    ' Vertical
    rows_to_fill = Range(range_rows).End(xlDown).Row
    range_fill = range_get & ":D" & CStr(rows_to_fill)
    Set SourceRange = Worksheets("Compliance").Range(range_get)
    Set fillRange = Worksheets("Compliance").Range(range_fill)
    SourceRange.AutoFill Destination := fillRange
    'By rows
    While count <= Range(range_rows).End(xlDown).Row
        range_get2 = "D" & CStr(count)
        range_fill = "D" & CStr(count) & ":BZ" & CStr(count)
        Set SourceRange = Worksheets("Compliance").Range(range_get2)
        Set fillRange = Worksheets("Compliance").Range(range_fill)
        SourceRange.AutoFill Destination:=fillRange
        count = count + 1
    Wend
    End Sub
    
'//////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub AuditsReport()

    Dim incorrect As Integer
    Dim range_col As Integer
    Dim range_rows As String
    Dim copy As String
    Dim paste As String
    Dim search_range As String
    'Getting the number of rows with data in the report
    incorrect = Workbooks("Audits").Worksheets("Sheet1").Range("A2").End(xlDown).row
    range_rows = Cells(incorrect, 5).Address
    range_col = 2
    'Getting all the values marked as "Incorrect" form one file and pasting them to a auxiliar one to review before sending it 
    While range_col < incorrect
        search_range = "AA" & CStr(range_col) & ":CB" & CStr(range_col)' The range to look up
        If (IsNumeric(Application.Match("Incorrect", Workbooks("Audits").Worksheets("Sheet1").Range(search_range), 0))) Then
            copy = "E" & CStr(range_col) 'Incorrect values
            paste = "C" & CStr(range_col + 4) 'Correct/Updated values
            Workbooks("auxiliar.xlsm").Worksheets("Compliance").Range(paste).Value = _
            Workbooks("Audits.xlsx").Worksheets("Sheet1").Range(copy).Value
        End If
        range_col = range_col + 1
    Wend
    
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////



