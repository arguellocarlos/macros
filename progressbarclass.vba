Option Explicit

Public WithEvents ProgressBarForm As UserForm

Public Sub InitializeProgressBar()
    Set ProgressBarForm = VBA.UserForms.Add("Forms.UserForm1")
    ProgressBarForm.Show
    DoEvents
End Sub

Public Sub UpdateProgressBar(currentStep As Integer, totalSteps As Integer)
    If Not ProgressBarForm Is Nothing Then
        Dim progressPercentage As Integer
        If currentStep <= totalSteps Then
            progressPercentage = Int((currentStep / totalSteps) * 100)
            ProgressBarForm.Controls("Label1").Caption = "Exporting emails to Excel: " & progressPercentage & "% complete"
            DoEvents
        End If
    End If
End Sub

Public Sub HideProgressBar()
    If Not ProgressBarForm Is Nothing Then
        ProgressBarForm.Hide
        Unload ProgressBarForm
        Set ProgressBarForm = Nothing
    End If
End Sub
