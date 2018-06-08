'VBA used for auto downloading the attachment using rules set up from outlook

Public Sub saveAttachtoDisk_TX(itm As Outlook.MailItem)
Dim objAtt As Outlook.Attachment
Dim saveFolder As String, RPath As String, RPathY As String, RPathM As String
Dim Y As String
Dim TodaysDate As String
Dim TodaySP() As String
Dim Mon As String
Dim TDay As String

'save to a defined place of your own
saveFolder = ""

     For Each objAtt In itm.Attachments
     'filter the file fomart wish to download
        If InStr(objAtt.FileName, "xls") <> 0 Or InStr(objAtt.FileName, "xlsx") <> 0 Or InStr(objAtt.FileName, "zip") <> 0 Then
          objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
          Set objAtt = Nothing
        End If
     Next
End Sub
