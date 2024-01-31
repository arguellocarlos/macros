Sub InsertTemplateContent()
    Dim templateFilePath As String
    Dim templateContent As String
    Dim currentItem As MailItem
    
    ' Set the path to your Outlook template file
    templateFilePath = "C:\Path\To\Your\Template.oft"
    
    ' Check if a mail item is currently selected
    If Application.ActiveExplorer.Selection.Count = 1 Then
        If TypeOf Application.ActiveExplorer.Selection.Item(1) Is MailItem Then
            Set currentItem = Application.ActiveExplorer.Selection.Item(1)
        Else
            MsgBox "Please select an email before running this macro.", vbExclamation
            Exit Sub
        End If
    Else
        ' If no mail item is selected, create a new mail item
        Set currentItem = Application.CreateItem(olMailItem)
    End If
    
    ' Check if the template file exists
    If Dir(templateFilePath) <> "" Then
        ' Read the content of the template file
        templateContent = ReadFileContent(templateFilePath)
        
        ' Insert the template content into the current message
        currentItem.HTMLBody = templateContent & currentItem.HTMLBody
    Else
        MsgBox "Template file not found.", vbExclamation
    End If
End Sub

Function ReadFileContent(filePath As String) As String
    Dim fileNumber As Integer
    Dim content As String
    
    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Read the content of the file
    content = Input$(LOF(fileNumber), fileNumber)
    
    ' Close the file
    Close fileNumber
    
    ' Return the file content
    ReadFileContent = content
End Function
