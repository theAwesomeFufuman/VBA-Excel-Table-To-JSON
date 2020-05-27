Attribute VB_Name = "ExcelToJSON"
'Declare variable for storing slected tables
Public SlctTbls()
'Declare variable for storing slected table worksheets
Public SlctSheets()
'Declare variable for storing numbers of tables in the workbook
Public TableCount As Integer
'Declare variable for the file name storage
Public jsonFile As String
Sub ExcelToJSON()
    'Turn off screen updates
    Application.ScreenUpdating = False
    'Declare variables for storing number of controls in the userform ExcelToJSONForm
    Dim FormCtrlCount As Integer
    FormCtrlCount = ExcelToJSONForm.Controls.Count
    'Declare variables for tables and sheets
    Dim Table As ListObject
    Dim Sheet As Worksheet
    'Declare array for storing table names
    Dim TableArray()
    'Declare variable for increasing the TableArray length
    Dim x As Integer
    x = 1
    'Declare variable for storing names of generated checkboxes
    Dim ChBxName As String
    'Loop through all Worksheets in the workbook
    For Each Sheet In Worksheets
        'Loop through all tables in the workbook and count them
        For Each Table In Sheet.ListObjects
            TableCount = TableCount + 1
            'Increase the TableArray length to be equal to the number of tables in the workbook
            ReDim Preserve TableArray(0 To x + TableCount)
            'Set the name of the current item in TableArray to be the same as the current table
            TableArray(TableCount) = Table.Name
            'Generate checkboxes in the userform ExcelToJSONForm
            ChBxName = "Table " & TableCount
            setcontrol = ExcelToJSONForm.Controls.Add("forms.checkbox.1", ChBxName, True)
        Next Table
    Next
    'Declare variable for looping through checkboxes
    Dim checkBox As Object
    'Loop through all checkboxes in the userform ExcelToJSONForm
    For Each checkBox In ExcelToJSONForm.Controls
        'Increase iteration number
        j = j + 1
        'If the iteration has gone past the number of controls that existed before the checkboxes were created...
        If j > FormCtrlCount Then
            'Set poition, caption to the respective table name, autosize
            With checkBox
                .Top = (j * 30) - 130
                .Left = 18
                .Caption = TableArray(j - FormCtrlCount)
                .AutoSize = True
            End With
        End If
    Next checkBox
    'Declare variable for user form height
    Dim UsrFormHeight As Integer
    UsrFormHeight = ((ExcelToJSONForm.Controls.Count) * 30)
    'Set the position of SubmitBtn and CancelBtn
    ExcelToJSONForm.SubmitBtn.Top = UsrFormHeight - 70
    ExcelToJSONForm.CancelBtn.Top = UsrFormHeight - 70
    'Show the userform ExcelToJSONForm
    With ExcelToJSONForm
        .Height = UsrFormHeight
        .Width = 400
        .Show
    End With
    'Wait until ExcelToJSONForm is hidden before proceeding
    Do While ExcelToJSONForm.Visible = True
    Loop
    'Declare variable for storing desired file location
    jsonFile = Application.GetSaveAsFilename(FileFilter:="JSON Files (*.json), *.json")
    If jsonFile <> "" Then
        'Open the file to start editing
        Open jsonFile For Output As #1
        'Declare variable for " character
        Dim strQuote As String
        strQuote = Chr$(34)
        'Declare variable for storing Workbook name without file extension
        Dim WBName As String
        WBName = strQuote & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name)) & strQuote
        'Print initial curly bracket, name of the workbook and an array for all tables in the workbook
        Print #1, "{"
        Print #1, strQuote & "Workbook" & strQuote & ": " & WBName & ","
        Print #1, strQuote & "Tables" & strQuote & ": " & "["
        'Select the first worksheet in the workbook
        Worksheets(1).Activate
        'Loop through all Worksheets in the workbook for the 2nd time
        For Each Sheet In Worksheets
            'Loop through all tables in the workbook for the 2nd time
            For Each Table In Sheet.ListObjects
                For i = 0 To UBound(SlctTbls)
                    If Table.Name = SlctTbls(i) Then
                        'Count how many tables has been looped through
                        TableCount = TableCount - 1
                        'Print initial curly bracket for the current table
                        Print #1, "{"
                        Print #1, strQuote & "TableName" & strQuote & ": " & strQuote & Table.Name & strQuote & ","
                        'Loop through all rows in the current table
                        For y = 1 To Table.ListRows.Count
                            Table.ListRows(y).Range.Select
                            'If the cell is empty, generate a name for the JSON Key
                            Dim tblRowKeyVal As String
                            'tblRowKeyVal = IIf(((ActiveCell.Value) = ""), WorksheetFunction.Concat(Table.Name, y), ActiveCell.Value.Replace(strQuote, "\" & strQuote))
                            tblRowKeyVal = IIf(((ActiveCell.Value) = ""), WorksheetFunction.Concat(Table.Name, y), Replace(CStr(ActiveCell.Value), strQuote, "\" & strQuote))
                            'Print the cell value of the first cell in the row as a JSON key followed by curly brackets
                            Print #1, strQuote & tblRowKeyVal & strQuote & ": {"
                            'Loop through all cells in current row, start with the second one
                            For j = 2 To Table.ListColumns.Count
                                'Select the second cell from the left in the row
                                ActiveCell.Offset(, 1).Activate
                                'Print the column header as key and cell contents as value
                                'Print #1, strQuote & Table.HeaderRowRange(j).Value & strQuote & ": " & strQuote & CStr(ActiveCell.Value.Replace(strQuote, "\" & strQuote)) & strQuote
                                Print #1, strQuote & Table.HeaderRowRange(j).Value & strQuote & ": " & strQuote & Replace(CStr(ActiveCell.Value), strQuote, "\" & strQuote) & strQuote
                                'If the cell is not the last iteration, then print a ","
                                If j < Table.ListColumns.Count Then
                                    Print #1, ","
                                End If
                            Next j
                            'Reselect the first cell of the row
                            ActiveCell.Offset(, (Table.ListColumns.Count * -1) + 1).Activate
                            'If the loop is not on the last iteration, put a "," after the ending curly bracket, otherwise, skip it
                            If y < Table.ListRows.Count Then
                                Print #1, "},"
                            Else
                                Print #1, "}"
                            End If
                        Next
                        'If the loop is not on the last iteration, put a "," after the ending curly bracket, otherwise, skip it
                        If TableCount > 0 Then
                            Print #1, "},"
                        Else
                            Print #1, "}"
                        End If
                    End If
                Next
            Next Table
            'If the loop has come to the last worksheet, activate the first one again
            If ActiveSheet.Index = Worksheets.Count Then
            Else
                Worksheets(ActiveSheet.Index + 1).Activate
            End If
        Next
        'Print closing square bracket and curly bracket
        Print #1, "]"
        Print #1, "}"
        'Close the file editing
        Close #1
        End
    Else
        End
    End If
    'Turn on screen updates
    Application.ScreenUpdating = True
    MsgBox ("The table(s) were successfully exported to " + jsonFile)
End Sub
