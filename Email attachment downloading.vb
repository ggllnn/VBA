'VBA used for downloading the attachment from selected mails to selected folders

Public Sub saveAttachtoPickedfolder()
Dim objAtt As Outlook.Attachment
Dim sel As Selection
Dim itm As Outlook.MailItem

Set sel = ActiveExplorer.Selection

saveFolder = ChoosePath_Folder()

For Each itm In sel

     For Each objAtt In itm.Attachments
          objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
          Set objAtt = Nothing
     Next
     
Next
End Sub

Public Function ChoosePath_Folder() As String
ChoosePath_Folder = ""
Dim excelapp As New Excel.Application
 Dim dlgOpen As office.FileDialog
  Set dlgOpen = excelapp.Application.FileDialog(msoFileDialogFolderPicker)
  With dlgOpen
          If .Show = -1 Then

                   ChoosePath_Folder = .SelectedItems(1)
          End If
  End With
  excelapp.Quit
  Set dlgOpen = Nothing
  Set excelapp = Nothing
End Function
