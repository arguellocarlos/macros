Option Explicit

Public WithEvents xlApp As Object
Public WithEvents xlWB As Object
Public WithEvents xlSheet As Object

Public Sub InitializeProgressBar()
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)
    
    ' Add a progress bar to the Excel sheet
    xlSheet.Shapes.AddFormControl(xlDialogButtonControl, Left:=10, Top:=10, Width:=150, Height:=15).Select
    xlApp.VBE.ActiveWindow.Visible = False
    xlApp.VBE.ActiveWindow.Visible = True
End Sub

Public Sub UpdateProgressBar(currentStep As Integer, totalSteps As Integer)
    If Not xlSheet Is Nothing Then
        Dim progressBar As Shape
        Set progressBar = xlSheet.Shapes("Button 1") ' Assuming the progress bar is the first button control added

        If currentStep <= totalSteps Then
            Dim progressPercentage As Integer
            progressPercentage = Int((currentStep / totalSteps) * 100)
            progressBar.ControlFormat.Value = progressPercentage
            xlApp.VBE.ActiveWindow.Visible = False
            xlApp.VBE.ActiveWindow.Visible = True
        End If
    End If
End Sub

Public Sub HideProgressBar()
    If Not xlApp Is Nothing Then
        xlWB.Close False
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlWB = Nothing
        Set xlApp = Nothing
    End If
End Sub
