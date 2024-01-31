Sub ImportTemplate()
    Dim objMail As Outlook.MailItem
    Dim objTemplate As Outlook.MailItem
    Dim objInspector As Outlook.Inspector

    Set objMail = Application.ActiveInspector.CurrentItem
    Set objTemplate = Application.CreateItemFromTemplate("C:\path\to\template.oft")
    Set objInspector = objMail.GetInspector

    With objTemplate
        .HTMLBody = Replace(.HTMLBody, "cid:", "file:///" & .Parent.Path & "/")
        objInspector.WordEditor.Range.FormattedText = .GetInspector.WordEditor.Range.FormattedText
    End With
End Sub
