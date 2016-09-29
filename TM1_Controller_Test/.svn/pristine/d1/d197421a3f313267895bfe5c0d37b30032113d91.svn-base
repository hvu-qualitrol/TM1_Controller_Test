Public Class Tests
    Public Shared Function YesNo(ByVal Question As String, ByVal Title As String) As Boolean
        Dim yesnoDialog As New YesNoDialog

        yesnoDialog.Question.Text = Question
        yesnoDialog.Text = Form1.Serial_Number_Entry.Text + ":  " + Title
        If Form1.USB1208LS_ComboBox.SelectedIndex < 0 Then
            yesnoDialog.MenuStrip1.BackColor = Control.DefaultBackColor
        Else
            yesnoDialog.MenuStrip1.BackColor = ColorTranslator.FromHtml(Module1.AppHexColorStr(Form1.USB1208LS_ComboBox.SelectedIndex))
        End If

        Return yesnoDialog.ShowDialog = DialogResult.OK
    End Function
End Class
