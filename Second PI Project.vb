Public aux As Integer
Public col1, col2, col3, col4, col5, col6, col7, col8 As Integer


Sub find_columns()
    
    'We set up the columns to use in the report 
    col1 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("Vulnerability Name").Column
    col2 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("Reference").Column
    col3 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("# Occurrence").Column
    col4 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("Severity").Column
    col5 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("False positive (Yes/ No)").Column
    col7 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("Recommended Solution(for genuine)").Column
    col8 = ThisWorkbook.Sheets(2).Range("A1:Z1").Find("Description").Column
End Sub

'Main function that will call the rest of them.
Sub Fill_issues()
    find_columns
    
    Dim MyFile, spl, F_priority, false_p As String
    Dim index As Integer
    Dim eval As Boolean
    Dim range1 As Range
    
	'******************************************************
	'To open de XML file
    MyFile = Application.GetOpenFilename()
    
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.Async = False
    xmlDoc.Load (MyFile)
	'******************************************************
	
    Set range1 = ThisWorkbook.Sheets(1).Range("L7")
    eval = already_opened(range1.Value, file_chosen(MyFile, ThisWorkbook.Sheets(1).Buttons("Button 2").Caption))
    
    If MyFile <> False And eval = False Then
        Set nodeXML_RS = xmlDoc.getElementsByTagName("ReportSection") 'Getting the 7 sections of the report
        '*******************************************************************************
        'To obtain the variables, -priority- & -analisys-
        '*******************************************************************************
        Set node_XML_re = nodeXML_RS(0).getElementsByTagName("Refinement")
        spl = node_XML_re(0).Text
        spliting = Split(spl, ":")
        spliting2 = Split(spliting(1), " AND ")
        F_priority = spliting(2)
        F_priority = IIf(F_priority = "critical", Replace(F_priority, "c", "C", 1, 1), F_priority)
        F_priority = IIf(F_priority = "high", Replace(F_priority, "h", "H", 1, 1), F_priority)
        F_priority = IIf(F_priority = "medium", Replace(F_priority, "m", "M", 1, 1), F_priority)
        F_priority = IIf(F_priority = "low", Replace(F_priority, "l", "L", 1, 1), F_priority)
        '****************************************************************************************
        'Going through the report section 
        For n = 0 To nodeXML_RS.Length - 1
            Set nodeXML_re = nodeXML_RS(n).getElementsByTagName("Title") 'Getting all the tittles 
            If nodeXML_re(0).Text = "Executive Summary" Then 'First part of the report, FALSE POSITIVES
                '*******************************************************************************
                'To fill the false positives -Vulnerability Name, Occurrence, Sevierity, False positive (Yes/ No):
                '*******************************************************************************
                If IsEmpty(ThisWorkbook.Sheets(2).Range("A2")) Then
                    index = 2
                Else
                    index = ThisWorkbook.Sheets(2).Range("A1").End(xlDown).row + 1
                End If
                index = Category_and_Count(nodeXML_RS, "Yes", index, n, F_priority)
            End If
			
            If nodeXML_re(0).Text = "Issue Count by Category" Then 'Second part of the report, GENUINES
                '*******************************************************************************
                'Filling genuines -Vulnerability Name, Occurrence, Sevierity, False positive (Yes/ No):
                '*******************************************************************************
                index = Category_and_Count(nodeXML_RS, "No", index, n, F_priority)
                
            End If
        Next
		'****************************************************************************************
		'Validating if the XML file was opened correctly, if it was, then you can run the actual macro, otherwise a message is displayed  
        If ThisWorkbook.Sheets(2).Range("M7").Value = 0 Then
            create_button
            ThisWorkbook.Sheets(2).Range("M7").Value = 1
        End If
        formating_cells
        range1.Value = range1.Value & file_chosen(MyFile, ThisWorkbook.Sheets(1).Buttons("Button 1").Caption)
        Else
            MsgBox ("La macro no se ejecutó pues no seleccionó ningún archivo o el archivo es incorrecto")
    End If
End Sub

'***********************************************************
'Creates the button that enables auditor to execute the Fill_references Sub
Sub create_button()
	
	Dim btn As Button
      Dim t As Range
      
        Set t = ThisWorkbook.Sheets(1).Range(Cells(7, 13), Cells(7, 13))
        Set btn = ThisWorkbook.Sheets(1).Buttons.Add(t.Left, t.Top, t.Width, t.Height)
        
        With btn
          .OnAction = "Fill_references"
          .Caption = "References"
          .name = "Button 1"
        End With
        
      Application.ScreenUpdating = True
End Sub
'***********************************************************

'***********************************************************
'Automatiaclly gives required format to the report cells
Sub formating_cells()
    With ThisWorkbook.Sheets(2).Range("A2:K1000").Borders()
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With ThisWorkbook.Sheets(2).Range("A2:K1000")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlJustify
    End With
        ThisWorkbook.Sheets(2).Columns("D").ColumnWidth = 60
        ThisWorkbook.Sheets(2).Columns("I:J").ColumnWidth = 60
        ThisWorkbook.Sheets(2).Columns("B:C").ColumnWidth = 30
        ThisWorkbook.Sheets(2).Columns("E:H").ColumnWidth = 15
        ThisWorkbook.Sheets(2).Rows("2:1000").RowHeight = 30
End Sub
'***********************************************************

'***********************************************************
'Leaves document as a template, ready for the next report
Sub clean_evry()
    ThisWorkbook.Sheets(2).Rows("2:1000").Value = ""
    On Error Resume Next
    ThisWorkbook.Sheets(1).Buttons("Button 1").Delete
    
    
    
    ThisWorkbook.Sheets(1).Range("L7").Value = ""
End Sub
'***********************************************************

'***********************************************************
'This is one of the most important Subs of the macro
'It is used to fill the corresponding cell with all the occurrences found of a certain issue references = occurrences = links
Sub Fill_references()
    find_columns
    
    Dim MyFile, spl, F_priority, false_p, name As String
    Dim index, vuln_row As Integer
    Dim eval As Boolean
    Dim range1 As Range
    
    F_priority = ""
    name = ""
    MyFile = Application.GetOpenFilename()
    Set range1 = ThisWorkbook.Sheets(1).Range("L7")
    '******************************************************
	'Open the XML document
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.Async = False
    xmlDoc.Load (MyFile)
    '******************************************************
    eval = already_opened(range1.Value, file_chosen(MyFile, ThisWorkbook.Sheets(1).Buttons("Button 1").Caption))
    If MyFile <> False And eval = False Then
    
        Set nodeXML_RS = xmlDoc.getElementsByTagName("ReportSection")
        aux = results_outline(nodeXML_RS)
        
        Set nodeXML_RS = nodeXML_RS(aux).getElementsByTagName("Chart")
        Set nodeXML_re = nodeXML_RS(0).getElementsByTagName("Issue")
		'******************************************************
        'Goes through all the issues on the XML document and gets their respective occurrences
		'This for false positives
        For i = 0 To nodeXML_re.Length - 1
            
            Set nodeXML_aux = nodeXML_re(i).getElementsByTagName("Friority")
            F_priority = nodeXML_aux(0).Text
            Set nodeXML_aux2 = nodeXML_re(i).getElementsByTagName("Category")
            name = nodeXML_aux2(0).Text
            
            vuln_row = find_vuln_book(name, "Yes", F_priority) ' this differentiates the next For loop
            
            Set nodeXML_aux3 = nodeXML_re(i).getElementsByTagName("FileName")
            
            If (vuln_row <> -1) Then
                ThisWorkbook.Sheets(2).Range(Cells(vuln_row, col2).Address).Value = nodeXML_aux3(0).Text
            End If
        Next
		'******************************************************
		'Goes through all the issues on the XML document and gets their respective occurrences
		'This for false positives
        For i = 0 To nodeXML_re.Length - 1
            
            Set nodeXML_aux = nodeXML_re(i).getElementsByTagName("Friority")
            F_priority = nodeXML_aux(0).Text
            Set nodeXML_aux2 = nodeXML_re(i).getElementsByTagName("Category")
            name = nodeXML_aux2(0).Text
            
            vuln_row = find_vuln_book(name, "No", F_priority)
            
            Set nodeXML_aux3 = nodeXML_re(i).getElementsByTagName("FileName")
            
            If (vuln_row <> -1) Then
                ThisWorkbook.Sheets(2).Range(Cells(vuln_row, col2).Address).Value = nodeXML_aux3(0).Text
            End If
        Next
		'******************************************************
        range1.Value = range1.Value & file_chosen(MyFile, ThisWorkbook.Sheets(1).Buttons("Button 1").Caption)
    Else
        MsgBox ("La macro no se ejecutó pues no seleccionó ningún archivo o el archivo es incorrecto")
    End If
End Sub
'***********************************************************


'***********************************************************
'Validating the XML file that's been chosen (Reports generated by the Fortify tool)
Function file_chosen(ByVal name As String, ByVal btn As String) As Integer
    file_chosen = -1
    
    If InStr(name, "References_U") And btn = "References" Then
        file_chosen = 4
    End If
    
    If InStr(name, "High_-_automating") Then
        file_chosen = 1
    End If
    If InStr(name, "Critical_-_automating") Then
        file_chosen = 5
    End If
    If InStr(name, "Low_-_automating") Then
        file_chosen = 3
    End If
    If InStr(name, "Medium_-_automating") Then
        file_chosen = 2
    End If
        
End Function
'***********************************************************

'****************************************************************
'This funtion evaluates if the file has been already open based on a value set on a cell
Function already_opened(ByVal numbers As String, ByVal chosen As Integer) As Boolean
        already_opened = False
    If InStr(numbers, CStr(chosen)) Or chosen = -1 Then
        already_opened = True
    End If
    If chosen = 4 And Not InStr(numbers, "1") And Not InStr(numbers, "2") And Not InStr(numbers, "3") And Not InStr(numbers, "5") Then
      already_opened = True
    End If
        
End Function
'****************************************************************

'****************************************************************
' Finds the vulnerability by its name, checking if it's false positive and the right priority.
Function find_vuln_book(ByVal vuln As String, ByVal falsep As String, ByVal priority As String) As Integer
    Dim index As Integer
    
    index = ThisWorkbook.Sheets(2).Range("A1").End(xlDown).row
    find_vuln_book = -1
    
    For i = 2 To index
        If (ThisWorkbook.Sheets(2).Range(Cells(i, col1).Address).Value = vuln) Then
            If (ThisWorkbook.Sheets(2).Range(Cells(i, col4).Address).Value = priority) Then
                If (ThisWorkbook.Sheets(2).Range(Cells(i, col5).Address).Value = falsep) Then
                    find_vuln_book = i
                    End If
            End If
        End If
    Next
End Function
'****************************************************************

'****************************************************************
'This funtion fills some of the values fo the vulnerabilities: category, priority, false positive, explanation and recomendation
Function Category_and_Count(ByVal nodeXML_RS As Variant, ByVal falsep As String, ByVal index As Integer, ByVal n As Integer, ByVal F_priority As String) As Integer
    Set nodeXML_re = nodeXML_RS(n).getElementsByTagName("Chart")
    Set nodeXML_cat = nodeXML_re(0).getElementsByTagName("groupTitle")
    Set nodeXML_count = nodeXML_re(0).getElementsByTagName("GroupingSection")
  
    For i = 0 To nodeXML_cat.Length - 1
        ThisWorkbook.Sheets(2).Range(Cells(index, 1).Address).Value = index - 1
        ThisWorkbook.Sheets(2).Range(Cells(index, col1).Address).Value = nodeXML_cat(i).Text 'CATEGORY
        ThisWorkbook.Sheets(2).Range(Cells(index, col3).Address).Value = nodeXML_count(i).getAttribute("count") 'issue count
        ThisWorkbook.Sheets(2).Range(Cells(index, col4).Address).Value = F_priority 'priority
        ThisWorkbook.Sheets(2).Range(Cells(index, col5).Address).Value = falsep ' False Positive (yes/no)
        Explanation_and_Recomendation nodeXML_RS, nodeXML_cat(i).Text, index, falsep ' Explanation and Recomendation
        index = index + 1
        
    Next
    Category_and_Count = index
End Function
'****************************************************************

'****************************************************************
'Getting the explanation and recomendation if it's genuine otherwise only the explanation 
Sub Explanation_and_Recomendation(ByVal nodeXML_RS As Variant, ByVal vuln As String, ByVal index As Integer, ByVal falsep As String)
    
    Set nodeXML_re = nodeXML_RS(aux).getElementsByTagName("Chart")
    Set nodeXML_cat = nodeXML_re(0).getElementsByTagName("groupTitle")
    Set NodeXML_section = nodeXML_re(0).getElementsByTagName("GroupingSection")
    Dim num_vuln As Integer
    
    num_vuln = vulnerability_pos(nodeXML_cat, vuln)
        Set nodeXML_value = NodeXML_section(num_vuln).getElementsByTagName("Value")
        If vuln = nodeXML_cat(num_vuln).Text Then
            If falsep = "Yes" Then
                ThisWorkbook.Sheets(2).Range(Cells(index, col8).Address).Value = Split_text(nodeXML_value(0).Text) 'Explanation
            End If
            If falsep = "No" Then
                ThisWorkbook.Sheets(2).Range(Cells(index, col8).Address).Value = Split_text(nodeXML_value(0).Text) 'Explanation
                ThisWorkbook.Sheets(2).Range(Cells(index, col7).Address).Value = Split_text(nodeXML_value(1).Text)	'Recomendation
            End If
        End If
		
End Sub
'****************************************************************

'****************************************************************
'Getting the index of the "Results Outline" field; this as a step to get the chart where the occurrences are located 
Function results_outline(ByVal nodeXML_RS As Variant) As Integer

    results_outline = -1
    For i = 0 To nodeXML_RS.Length - 1
        Set nodeXML_re = nodeXML_RS(i).getElementsByTagName("Title")
        If nodeXML_re(0).Text = "Results Outline" Then
            results_outline = i
        End If
    Next
    
End Function
'****************************************************************

'****************************************************************
'Getting the vulnerability position in the XML document based on its name
Function vulnerability_pos(ByVal nodeXML_cat As Variant, ByVal vuln As String) As Integer

    vulnerability_pos = -1
    For i = 0 To nodeXML_cat.Length - 1
         If vuln = nodeXML_cat(i).Text Then
            vulnerability_pos = i
        End If
    Next
    
End Function
'****************************************************************

'****************************************************************
'Automating the part of leaving just what's necesary on the explanation
Function Split_text(ByVal explanation As String) As String
    Dim index, count As Integer
    
    index = 0
    count = 0
    tres = Split(explanation, "Example ")
    
    If (tres(0) = explanation) Then
        tres = Split(explanation, "Example:")
    End If
    tres = Split(tres(0), vbLf)
    explanation = ""
    While (index < UBound(tres) + 1 And index < 8)
        count = count + Len(tres(index))
        If count > 1350 Then
            GoTo lab1
        Else
            explanation = explanation & tres(index) & vbLf
            index = index + 1
        End If
    Wend
lab1:
    Split_text = explanation
End Function
'****************************************************************