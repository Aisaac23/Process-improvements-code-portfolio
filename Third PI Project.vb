Sub CopyStepsToClipboard()

Dim shp As Shape
Dim str, curStep As String
Dim totalPages, nSteps, pageNumber, stepNumber, page, pageSteps, totsteps As Integer
Dim last As Boolean
Dim mainArray()

totalPages = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
nSteps = getTotalsteps()
stepNumber = 1
pageSteps = 0
totsteps = 0
last = False
    page = 1 
    While (Not last Or page <= totalPages) And totsteps < nSteps
        curStep = getNextStep(page, stepNumber)
        
        If curStep = "500" Then 'found but not numeric step
            str = str + "**not_numeric" + vbLf
        ElseIf getNextStep(page, stepNumber) = "404" And pageSteps >= stepsByPage(page) Then'not found and there are no more steps left in the page
            page = page + 1
            pageSteps = 1
        ElseIf getNextStep(page, stepNumber) = "404" And pageSteps < stepsByPage(page) Then'not found but there are steps left in the page
            str = str & stepNumber & "." + vbCrLf
            stepNumber = stepNumber + 1
        Else ' step successfully found 
            str = str + curStep + vbLf
            last = False
            stepNumber = stepNumber + 1
            pageSteps = pageSteps + 1
            totsteps = totsteps + 1
        End If
        'validating if it's the last step on the vulnerability
        If isLast(getNextStep(page, stepNumber)) Then
            str = str + (getNextStep(page, stepNumber)) + vbLf + vbCrLf
            totsteps = totsteps + 1
            stepNumber = 1
            last = True
        End If
        
        
    Wend

'Put the string with all the steps in CLIPBOARD
With New MSForms.DataObject
    .SetText str
    .PutInClipboard
End With

End Sub

'********************************************************************************
'Returns the steps by page based on all the shapes which have text inside of them that are inside of the given page 
Function stepsByPage(ByVal pageNo As Integer) As Integer

	Dim pn, count As Integer
	count = 0

	For Each shp In ActiveDocument.Shapes
		If Len(shp.TextFrame.TextRange.Text) > 1 Then
			
			pn = getPageNo(shp)
			If pn = pageNo Then
				count = count + 1
			End If
			
		End If
	Next

	stepsByPage = count

End Function
'********************************************************************************

'********************************************************************************
'Returns the  total steps in the whole document based on all the shapes which have text inside of them
Function getTotalsteps() As Integer

Dim pn, count As Integer
count = 0

For Each shp In ActiveDocument.Shapes
    If Len(shp.TextFrame.TextRange.Text) > 1 Then
            count = count + 1
    End If
Next

getTotalsteps = count

End Function
'********************************************************************************

'********************************************************************************
'Returns true if the step contains any of the specified strings of else otherwise; indicating if it's the last step of the vulnerability
Function isLast(ByVal step As String) As Boolean

    If InStr(1, step, " of the vulnerability") > 0 Or InStr(1, step, " this as proof") > 0 Then
        isLast = True
    Else
        isLast = False
    End If

End Function
'********************************************************************************

'********************************************************************************
'Returns a string with the step given as NextStep verifying that it's located in the correct page
Function getNextStep(ByVal pageNo As Integer, ByVal nextStep As Integer) As String

Dim str As String
Dim count, step, pn As Integer
count = 0
str = "404"' not found

For Each shp In ActiveDocument.Shapes
    If Len(shp.TextFrame.TextRange.Text) > 1 Then
    
        pn = getPageNo(shp)
        step = getStepNo(shp.TextFrame.TextRange.Text)
    
        If (step = nextStep And pn = pageNo) Then
            str = shp.TextFrame.TextRange.Text
        End If
        
        If step = -1 Then
            str = "500" 'there was a step but the first one or two characters were not numeric 
        End If
    End If
Next

getNextStep = str

End Function
'********************************************************************************

'********************************************************************************
'Returns the page number based on a given shape contained in that page 
Function getPageNo(ByVal nStepBox As Shape) As Integer

nStepBox.Select
pageNumber = Selection.Information(wdActiveEndPageNumber)

getPageNo = pageNumber

End Function
'********************************************************************************

'********************************************************************************
'Returns the step number based on a given string and using its first or its first and the second characters 
Function getStepNo(ByVal step As String) As Integer

Dim ch As String
Dim ch2 As String
Dim stepNo As Integer

ch = Mid(step, 1, 1)
ch2 = Mid(step, 2, 1)

If IsNumeric(ch) And IsNumeric(ch2) Then 'for to digit step number
    stepNo = CInt(ch + ch2)
ElseIf IsNumeric(ch) Then ' for one digit step number
    stepNo = CInt(ch)
Else
    stepNo = -1
End If

getStepNo = stepNo

End Function
'********************************************************************************


