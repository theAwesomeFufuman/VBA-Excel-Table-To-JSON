Attribute VB_Name = "ExcelToJSON"
Option Explicit

'Stores the names of tables that the user wants to convert to JSON.
'The names are added to the array during the function SubmitBtn_Click() in ExcelToJSONForm.
Public usrSlctdTblsNameArray() As String

'Variables for iterating through loops
Public i As Integer, j As Integer, k As Integer

Public outputFileFQPN As String

Public numIndentationSpaces As Integer
Public currentIndentation As Integer

Sub ExcelToJSON()
    
    numIndentationSpaces = 4
    
    Application.ScreenUpdating = False
    
    Dim originallySelectedSheet As String: originallySelectedSheet = ActiveSheet.Name
    
    Dim FALSEInLocalLang As String: FALSEInLocalLang = getLocalTranslationOfFALSE()
    
    Dim table As ListObject, sheet As Worksheet
    
    Dim numFormCtrlsNotCountingCheckBoxes As Integer: numFormCtrlsNotCountingCheckBoxes = ExcelToJSONForm.Controls.Count
    
    Dim tableNamesInCurrentWorkbook()
    
    Dim tblNameToChBxName As String
    
    'Loop through all of the tables in the workbook and generate checkboxes with corresponding names.
    'The checkboxes lets the user select which tables to export to JSON in ExcelToJSONForm.
    For Each sheet In Worksheets
        For Each table In sheet.ListObjects
            i = i + 1
            
            ReDim Preserve tableNamesInCurrentWorkbook(0 To i + 1)
            tableNamesInCurrentWorkbook(i) = table.Name
            
            tblNameToChBxName = table.Name & "ChBx"
            Dim addNewCheckbox As Boolean: addNewCheckbox = ExcelToJSONForm.Controls.Add("forms.checkbox.1", tblNameToChBxName, True)
            
        Next table
    Next
    
    i = 0
    Dim userFormControl As Object
    For Each userFormControl In ExcelToJSONForm.Controls
        i = i + 1
        
        If i > numFormCtrlsNotCountingCheckBoxes Then
            With userFormControl
                .Top = (i * 30) - 80
                .Left = 108
                .Caption = tableNamesInCurrentWorkbook(i - numFormCtrlsNotCountingCheckBoxes)
                .AutoSize = True
            End With
        End If
        
    Next userFormControl
    
    Dim UsrFormWindowHeight As Integer: UsrFormWindowHeight = ExcelToJSONForm.Controls.Count * 30
    
    ExcelToJSONForm.SubmitBtn.Top = UsrFormWindowHeight - 60
    ExcelToJSONForm.CancelBtn.Top = UsrFormWindowHeight - 60
    
    With ExcelToJSONForm
        .Height = UsrFormWindowHeight
        .Width = 400
        .Show
    End With
    
    Do While ExcelToJSONForm.Visible = True
    Loop
    
    outputFileFQPN = Application.GetSaveAsFilename(FileFilter:="JSON Files (*.json), *.json")
    If outputFileFQPN <> "" And outputFileFQPN <> FALSEInLocalLang Then
        
        Open outputFileFQPN For Output As #1
        
        'Print the opening bracket for the JSON object as well as a key/value for a string of the current file name and
        'the JSON key and opening bracket for the object representing all worksheets in the workbook
        Print #1, "{"
        Print #1, createJsonKeyValuePair(numIndentationSpaces * 1, "Source file name", ActiveWorkbook.Name, True)
        Print #1, createJsonKeyValuePair(numIndentationSpaces * 1, "Worksheets", "{", False)
        
        Worksheets(1).Activate
        
        Dim numTablesInSheet As Integer
        
        For Each sheet In Worksheets
            numTablesInSheet = 0
            
            'Print the JSON key and opening bracket for each object representing a worksheet
            Print #1, createJsonKeyValuePair(numIndentationSpaces * 2, sheet.Name, "{", False)
            
            'Print the JSON key and opening bracket for the "Tables" object inside each sheet object
            Print #1, createJsonKeyValuePair(numIndentationSpaces * 3, "Tables", "{", False)
            
            Dim printCommaUnlessLastIteration As Boolean
            
            For Each table In sheet.ListObjects
                numTablesInSheet = numTablesInSheet + 1
                
                For i = 0 To UBound(usrSlctdTblsNameArray)
                    If table.Name = usrSlctdTblsNameArray(i) Then
                    
                        'Print the JSON key and opening bracket for the object representing each table inside the "Tables" object
                        Print #1, createJsonKeyValuePair(numIndentationSpaces * 4, table.Name, "{", False)
                        
                        For j = 1 To table.ListRows.Count
                            table.ListRows(j).Range.Select
                            
                            Dim tableRowIndexCellValue As String: tableRowIndexCellValue = ActiveCell.Value
                            Dim tableNamePlusIterationNumber As String: tableNamePlusIterationNumber = WorksheetFunction.Concat(table.Name, j)
                            Dim tableRowKey As String: tableRowKey = IIf((tableRowIndexCellValue = ""), tableNamePlusIterationNumber, tableRowIndexCellValue)
                        
                            Print #1, createJsonKeyValuePair(numIndentationSpaces * 5, tableRowKey, "{", False)
                            
                            'Loop through all cells in the current row, starting with the 2nd cell from the left.
                            'The loop starts with the 2nd cell from the left because the 1st cell from the left is
                            'converted to a JSON Key for the rest of the cells in the same table row and the cells in the same row are converted to JSON Values.
                            For k = 2 To table.ListColumns.Count
                                ActiveCell.Offset(, 1).Activate
                                
                                printCommaUnlessLastIteration = k < table.ListColumns.Count
                                Print #1, createJsonKeyValuePair(numIndentationSpaces * 6, table.HeaderRowRange(k).Value, ActiveCell.Value, printCommaUnlessLastIteration)
                                
                            Next k
                            
                            'Reselect the index cell of the current row
                            ActiveCell.Offset(, (table.ListColumns.Count * -1) + 1).Activate
                            
                            printCommaUnlessLastIteration = j < table.ListRows.Count
                            Print #1, createJsonClosingBracket((numIndentationSpaces * 5), printCommaUnlessLastIteration)
                            
                        Next j
                        
                        printCommaUnlessLastIteration = numTablesInSheet < sheet.ListObjects.Count
                        Print #1, createJsonClosingBracket((numIndentationSpaces * 4), printCommaUnlessLastIteration)
                        
                    End If
                Next
            Next table
            
            Print #1, createJsonClosingBracket((numIndentationSpaces * 3), False)
            
            printCommaUnlessLastIteration = sheet.Index < Worksheets.Count
            
            Print #1, createJsonClosingBracket((numIndentationSpaces * 2), printCommaUnlessLastIteration)
            
            'If the loop is not on its last iteration, select the next sheet in the workbook
            If ActiveSheet.Index < Worksheets.Count Then
                Worksheets(ActiveSheet.Index + 1).Activate
            End If
        Next sheet
        
        Sheets(originallySelectedSheet).Select
        
        'Print closing square bracket and curly bracket
        Print #1, createJsonClosingBracket((numIndentationSpaces * 1), False)
        Print #1, createJsonClosingBracket((numIndentationSpaces * 0), False)
        
        Close #1
        End
    Else
        End
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Function createJsonKeyValuePair(numSpacesToIndent As Integer, keyString As String, valueString As String, showComma As Boolean) As String
    Dim output As String: output = ""
    
    For i = 1 To numSpacesToIndent
        output = output & " "
    Next i
    
    output = output & Chr$(34) & keyString & Chr$(34)
    output = output & ": "
    output = IIf(valueString = "{", output & "{", output & Chr$(34) & valueString & Chr$(34))
    output = IIf(showComma, output & ",", output)
    
    createJsonKeyValuePair = output
End Function

Public Function createJsonClosingBracket(numSpacesToIndent As Integer, showComma As Boolean) As String
    Dim output As String: output = ""
    
    For i = 1 To numSpacesToIndent
        output = output & " "
    Next i
    
    output = output & "}"
    output = IIf(showComma, output & ",", output)
    
    createJsonClosingBracket = output
End Function

Public Function getLocalTranslationOfFALSE()
    'The last possible cell "XFD1048576" is selected because of its low risk of containing sensitive iformation and thus
    'can be used to paste the local translation of "false"
    Dim originalValue As Variant: originalValue = Range("XFD1048576").Value
    
    Range("XFD1048576").Value = False
    Dim localTranslation As String
    localTranslation = Range("XFD1048576").Value
    
    Range("XFD1048576").Value = originalValue
    
    getLocalTranslationOfFALSE = localTranslation
End Function
