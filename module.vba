Option Explicit

Sub ExportEmailsToExcel()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim exportStartDate As Date
    Dim exportEndDate As Date
    Dim filterKeyword As String
    Dim totalEmails As Integer
    Dim exportedEmails As Integer
    Dim mailItem As Outlook.MailItem
    Dim excelHandler As ProgressBarClass
    Set excelHandler = New ProgressBarClass

    ' Prompt the user for date range and keyword
    exportStartDate = InputBox("Enter the start date for the export (MM/DD/YYYY):", "Start Date")
    exportEndDate = InputBox("Enter the end date for the export (MM/DD/YYYY):", "End Date")
    filterKeyword = InputBox("Enter the keyword to filter emails (e.g., INC):", "Filter Keyword")

    ' Prompt the user to select the folder
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olApp.Session.PickFolder

    ' Check if a folder is selected
    If olFolder Is Nothing Then
        MsgBox "No folder selected. Export canceled.", vbExclamation
        Exit Sub
    End If

    ' Count the total number of eligible emails
    totalEmails = 0
    For Each olItem In olFolder.Items
        If TypeOf olItem Is Outlook.MailItem Then
            Set mailItem = olItem
            If mailItem.ReceivedTime >= exportStartDate And mailItem.ReceivedTime <= exportEndDate And InStr(1, mailItem.Subject, filterKeyword, vbTextCompare) > 0 Then
                totalEmails = totalEmails + 1
            End If
        End If
    Next olItem

    ' Initialize Excel
    excelHandler.InitializeExcel

    ' Create Excel application and workbook
    Set xlApp = excelHandler.xlApp
    Set xlWB = excelHandler.xlWB
    Set xlSheet = excelHandler.xlSheet

    ' Set column headers in Excel
    xlSheet.Cells(1, 1).Value = "Sender Email"
    xlSheet.Cells(1, 2).Value = "Sender Name"
    xlSheet.Cells(1, 3).Value = "Received Date"
    xlSheet.Cells(1, 4).Value = "Subject"

    ' Loop through the emails in the specified folder
    Dim rowNum As Integer
    rowNum = 2 ' Start from row 2 to leave space for headers

    For Each olItem In olFolder.Items
        If TypeOf olItem Is Outlook.MailItem Then
            Set mailItem = olItem

            ' Check if the email meets the criteria
            If mailItem.ReceivedTime >= exportStartDate And mailItem.ReceivedTime <= exportEndDate And InStr(1, mailItem.Subject, filterKeyword, vbTextCompare) > 0 Then
                ' Export data to Excel
                xlSheet.Cells(rowNum, 1).Value = mailItem.SenderEmailAddress
                xlSheet.Cells(rowNum, 2).Value = mailItem.SenderName
                xlSheet.Cells(rowNum, 3).Value = mailItem.ReceivedTime
                xlSheet.Cells(rowNum, 4).Value = mailItem.Subject

                rowNum = rowNum + 1
                exportedEmails = exportedEmails + 1
            End If
        End If
    Next olItem

    ' Save the Excel file (you can modify the file path)
    xlWB.SaveAs "C:\Path\To\Your\ExportedEmails.xlsx"

    ' Close Excel without saving changes to the template
    xlWB.Close False
    xlApp.Quit

    ' Release object references
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set olItem = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing

    ' Hide Excel
    excelHandler.HideExcel

    MsgBox "Export completed successfully!", vbInformation
End Sub
