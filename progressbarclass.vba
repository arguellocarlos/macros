Option Explicit

Public WithEvents xlApp As Object
Public WithEvents xlWB As Object
Public WithEvents xlSheet As Object

Public Sub InitializeExcel()
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)
End Sub

Public Sub HideExcel()
    If Not xlApp Is Nothing Then
        xlWB.Close False
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlWB = Nothing
        Set xlApp = Nothing
    End If
End Sub
