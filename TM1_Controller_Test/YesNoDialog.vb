Public Class YesNoDialog

    Private Sub Yes_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = DialogResult.OK
        Me.Dispose()
    End Sub

    Private Sub No_Click(sender As System.Object, e As System.EventArgs) Handles No.Click
        Me.DialogResult = DialogResult.No
    End Sub
End Class