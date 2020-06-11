VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelToJSONForm 
   Caption         =   "Select table(s)"
   ClientHeight    =   2890
   ClientLeft      =   -950
   ClientTop       =   -4950
   ClientWidth     =   2480
   OleObjectBlob   =   "ExcelToJSONForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelToJSONForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare variable for looping through User Form Controls
Public userFormControl As Object
Private Sub CancelBtn_Click()
    End
End Sub
Private Sub SelectAll_Click()
    For Each userFormControl In ExcelToJSONForm.Controls
        If TypeName(userFormControl) = "CheckBox" Then
            userFormControl.Value = True
        End If
    Next userFormControl
End Sub
Private Sub SelectNone_Click()
    For Each userFormControl In ExcelToJSONForm.Controls
        If TypeName(userFormControl) = "CheckBox" Then
            userFormControl.Value = False
        End If
    Next userFormControl
End Sub
Private Sub SubmitBtn_Click()
    Dim numCheckedBoxes As Integer

    For Each userFormControl In ExcelToJSONForm.Controls
        If TypeName(userFormControl) = "CheckBox" Then
            If userFormControl = True Then
                numCheckedBoxes = numCheckedBoxes + 1
            End If
        End If
    Next userFormControl
    
    If numCheckedBoxes = 0 Then
        MsgBox "Please select one or more tables before proceeding"
    Else
        
        j = 0
        
        ReDim Preserve usrSlctdTblsNameArray(0 To numCheckedBoxes)
        For Each userFormControl In ExcelToJSONForm.Controls
            If TypeName(userFormControl) = "CheckBox" Then
                If userFormControl = True Then
                    j = j + 1
                    usrSlctdTblsNameArray(j) = userFormControl.Caption
                End If
            End If
        Next userFormControl
        Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        End
    End If
End Sub
