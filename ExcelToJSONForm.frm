VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelToJSONForm 
   Caption         =   "Select table(s)"
   ClientHeight    =   2710
   ClientLeft      =   -230
   ClientTop       =   -1050
   ClientWidth     =   5860
   OleObjectBlob   =   "ExcelToJSONForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelToJSONForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare variable for looping through User Form Controls
Public usrFrmCtrl As Object

Private Sub CancelBtn_Click()
    End
End Sub
Private Sub SelectAll_Click()
    'Loop through all controls in the userform
    For Each usrFrmCtrl In ExcelToJSONForm.Controls
        'Check if the control is a checkbox and if it is visible
        If TypeName(usrFrmCtrl) = "CheckBox" And usrFrmCtrl.Visible = True Then
            'Set all checkboxes to the same value as TrueBox(True)
            usrFrmCtrl.Value = Me.TrueBox.Value
        End If
    Next usrFrmCtrl
End Sub
Private Sub SelectNone_Click()
    'Loop through all controls in the userform
    For Each usrFrmCtrl In ExcelToJSONForm.Controls
        'Check if the control is a checkbox and if it is visible
        If TypeName(usrFrmCtrl) = "CheckBox" And usrFrmCtrl.Visible = True Then
            'Set all checkboxes to the same value as FalseBox(False)
            usrFrmCtrl.Value = Me.FalseBox.Value
        End If
    Next usrFrmCtrl
End Sub
Private Sub SubmitBtn_Click()
    'Declare variable for storing number of checked checkboxes
    Dim Checked As Integer
    'Loop through all controls in the userform
    For Each usrFrmCtrl In ExcelToJSONForm.Controls
        'Check if the control is a checkbox and if it is visible
        If TypeName(usrFrmCtrl) = "CheckBox" And usrFrmCtrl.Visible = True Then
            'If a checkbox is checked, increase the number of checked checkboxes
            If usrFrmCtrl = True Then
                Checked = Checked + 1
            End If
        End If
    Next usrFrmCtrl
    'If no checkboxes are checked...
    If Checked = 0 Then
        'Display an error message
        MsgBox "Please select one or more tables before proceeding"
    Else
        'Loop through all controls in the userform
        For Each usrFrmCtrl In ExcelToJSONForm.Controls
            'Check if the control is a checkbox and if it is visible
            If TypeName(usrFrmCtrl) = "CheckBox" And usrFrmCtrl.Visible = True Then
                'If a checkbox is checked, increase the number of checked checkboxes
                If usrFrmCtrl = True Then
                    j = j + 1
                    'Increase the SlctTbls array length to be equal to the slected number of tables
                    ReDim Preserve SlctTbls(0 To j + 1)
                    'Increase TableCount to be equal to the number of selected tables
                    TableCount = UBound(SlctTbls) - 1
                    'Add the table names to SlctTbls
                    SlctTbls(j) = usrFrmCtrl.Caption
                End If
            End If
        Next usrFrmCtrl
        'Hide the form
        Me.Hide
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        End
    End If
End Sub
