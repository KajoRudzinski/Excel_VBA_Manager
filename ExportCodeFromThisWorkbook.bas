Attribute VB_Name = "ExportCodeFromThisWorkbook"
Option Explicit

'   Author: Kajo Rudzinski
'
'   Inspired by:
'       https://www.rondebruin.nl/
'       https://www.youtube.com/user/WiseOwlTutorials

Private Const MsgGetExportFolderReason As String = _
    "You will now be asked to select the folder " & _
    "where you would like to have your code exported to."
    
Private Const MsgFolderNotSelected As String = _
    "The folder has not been selected."
    
Private Const MsgVBAProtectedWorkbook As String = _
    "The VBA in this workbook is protected," & _
    "so it's not possible to export the code"
    
Private Const MsgEndWithoutExport As String = _
    "The procedure ends now, no export has been made."
    
Private Const MsgExportSuccessful As String = _
    "The export was successful." & _
 _
    vbNewLine & vbNewLine & _
 _
    "Have a good day :-)"
    
Private Const MsgUserAgreementQuestion As String = _
    "Thank you for using this module." & _
 _
    vbNewLine & vbNewLine & _
 _
    "Would you like to export VBA code from this workbook?" & _
 _
    vbNewLine & vbNewLine & _
 _
    "NOTE: All previously existing exported code " & _
    "with the same name in the chosen export folder " & _
    "will be overwritten."

Private FileToExportFrom As Workbook
Private ExportFolder As String
Private NewFileDialog As FileDialog
Private SelectionResult As Boolean
Private CodeComponent As VBIDE.VBComponent
Private UserAgreement As Byte


    Sub ExportCodeFromThisWorkbook()
    
    '###########################################
    '
    '### Call this procedure to execute ###
    '
    '###########################################
    
    If UserAgreedToProceed Then
        Call CheckIfCodeIsNotProtected
        ExportFolder = GetExportFolder & "\"
        Set FileToExportFrom = ActiveWorkbook
        Call ExportCodeComponents(ExportFolder, FileToExportFrom)
        MsgBox MsgExportSuccessful, vbInformation, "Success"
        Else
            MsgBox MsgEndWithoutExport
    End If
    End Sub
    '############################################


Private Function UserAgreedToProceed() As Boolean

    UserAgreement = MsgBox(MsgUserAgreementQuestion, _
    vbYesNo + vbQuestion, "Agreement to proceed")
    
    If UserAgreement = vbYes Then
        UserAgreedToProceed = True
    End If

End Function


Private Sub ExportCodeComponents _
(ExportFolder As String, FileToExportFrom As Workbook)

    For Each CodeComponent In FileToExportFrom.VBProject.VBComponents
        CodeComponent.Export _
        GetCodeComponentExportPath(ExportFolder, CodeComponent)
        
        Debug.Print Now & _
        GetCodeComponentExportPath(ExportFolder, CodeComponent) & " exported"
    Next CodeComponent

End Sub

Private Function GetCodeComponentExportPath _
(ExportFolder As String, CodeComponent As VBIDE.VBComponent) As String

    GetCodeComponentExportPath = _
        ExportFolder & CodeComponent.Name & _
        GetCodeComponentExtension(CodeComponent)

End Function

Private Function GetCodeComponentExtension _
(CodeComponent As VBIDE.VBComponent) As String

    Select Case CodeComponent.Type
        Case vbext_ct_ClassModule
            GetCodeComponentExtension = ".cls"
        Case vbext_ct_MSForm
            GetCodeComponentExtension = ".frm"
        Case vbext_ct_StdModule
            GetCodeComponentExtension = ".bas"
        Case vbext_ct_Document
            GetCodeComponentExtension = ".cls"
    End Select

End Function


Private Sub CheckIfCodeIsNotProtected()
    If CodeIsProtected Then
        MsgBox _
            MsgVBAProtectedWorkbook & _
            vbNewLine & vbNewLine & _
            MsgEndWithoutExport
        End
    End If
End Sub

Private Function CodeIsProtected() As Boolean
    If ActiveWorkbook.VBProject.Protection = vbext_pp_locked Then
        CodeIsProtected = True
    End If
End Function

Private Function GetExportFolder() As String
    
    Set NewFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    NewFileDialog.Title = "Select the folder to export to..."
    MsgBox MsgGetExportFolderReason
    
    Call TestFolderSelection(NewFileDialog.Show)
    
    GetExportFolder = NewFileDialog.SelectedItems(1)

End Function

Private Sub TestFolderSelection(SelectionToTest As Boolean)
    If SelectionToTest = False Then
        MsgBox _
            MsgFolderNotSelected & _
            vbNewLine & vbNewLine & _
            MsgEndWithoutExport
        End
    End If
End Sub
