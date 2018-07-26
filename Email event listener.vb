Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    Set Msg = item
    'do something here
    '*****************
    '
    '*****************
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub
