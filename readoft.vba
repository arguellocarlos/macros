Option Explicit

Sub InsertTemplateIntoEmail()
    Dim olApp As Outlook.Application
    Dim olInspector As Outlook.Inspector
    Dim olMail As Outlook.MailItem
    Dim olTemplate As Outlook.MailItem
    Dim templatePath As String

    ' Prompt the user to select an Outlook template (.oft file)
    templatePath = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderTemplates).FolderPath
    templatePath = Application.GetNamespace("MAPI").PickFolder.Items.Item("Outlook Template").FolderPath
    If templatePath = "" Then
        MsgBox "No template selected. Operation canceled.", vbExclamation
        Exit Sub
    End If

    ' Create a new Outlook Application
    Set olApp = CreateObject("Outlook.Application")

    ' Create a new mail item (you can use the active item if desired)
    Set olInspector = olApp.ActiveInspector
    If Not olInspector Is Nothing Then
        If olInspector.CurrentItem.Class = olMail Then
            Set olMail = olInspector.CurrentItem
        End If
    End If

    ' Check if a mail item is selected
    If olMail Is Nothing Then
        MsgBox "No mail item selected. Operation canceled.", vbExclamation
        Set olApp = Nothing
        Exit Sub
    End If

    ' Open the template
    Set olTemplate = Application.CreateItemFromTemplate(templatePath)

    ' Copy the content from the template to the current/new email
    CopyTemplateContent olTemplate, olMail

    ' Clean up
    Set olTemplate = Nothing
    Set olMail = Nothing
    Set olInspector = Nothing
    Set olApp = Nothing

    MsgBox "Template inserted successfully!", vbInformation
End Sub

Sub CopyTemplateContent(sourceMail As Outlook.MailItem, destinationMail As Outlook.MailItem)
    ' Copy the content from the source template to the destination mail item
    Dim sourceWordDoc As Object
    Dim destinationWordDoc As Object
    Dim wdPasteDefault As Integer

    ' Word constants (late binding)
    Const wdPasteDefaultConst As Integer = 0 ' Use wdPasteDefault constant value if Word is not referenced

    ' Get the Word editor for both source and destination mails
    Set sourceWordDoc = sourceMail.GetInspector.WordEditor
    Set destinationWordDoc = destinationMail.GetInspector.WordEditor

    ' Copy content and formatting
    sourceWordDoc.Range.Copy
    destinationWordDoc.Range.PasteSpecial Format:="HTML", Link:=False, Placement:=wdPasteDefaultConst

    ' Clean up
    Set sourceWordDoc = Nothing
    Set destinationWordDoc = Nothing
End Sub
