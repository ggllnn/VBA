'Microsoft Outlook Object Library will be required

Sub Email_Send()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        'display the mail created so that you can see
        .display
        'To email address
        .To = ""
        'CC email address
        .CC = ""
        'BCC email address
        .BCC = ""
        .Subject = ""
        .HTMLBody = ""
        .Attachments.add ActiveWorkbook.FullName
        'You can add other files also like this below
        '.Attachments.add (FilePath & "\" & FileName)
        .send

    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
