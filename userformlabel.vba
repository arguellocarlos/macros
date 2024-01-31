Option Explicit

Public Sub ShowProgressBar(totalSteps As Integer)
    Me.Show
    Me.Label1.Caption = "Exporting emails to Excel: 0% complete"
    Me.Label1.Tag = totalSteps
    DoEvents
End Sub

Public Sub UpdateProgressBar(currentStep As Integer)
    Dim progressPercentage As Integer
    If currentStep <= Me.Label1.Tag Then
        progressPercentage = Int((currentStep / Me.Label1.Tag) * 100)
        Me.Label1.Caption = "Exporting emails to Excel: " & progressPercentage & "% complete"
        DoEvents
    End If
End Sub

Public Sub HideProgressBar()
    Me.Hide
End Sub