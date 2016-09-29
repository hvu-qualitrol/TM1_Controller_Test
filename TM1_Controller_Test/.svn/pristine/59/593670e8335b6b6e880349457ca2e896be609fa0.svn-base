Partial Class Tests
    Public Shared Function Test_FLASH() As Boolean
        Dim JL As New Jlink
        Dim Version As String
        'Dim DR As DialogResult

        Version = Products(Product)("FW_VERSION")

        If Not YesNo("Power up the controller board", "Power up?") Then
            Form1.AppendText("Operator indicated problem powering up.")
            Return False
        End If

        'DR = MessageBox.Show("Power up the controller board", "Power up?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem powering up.")
        '    ' LedTimer.Stop()
        '    Return False
        'End If

        Form1.AppendText("Verifying Flash", True)
        If JL.Verify(Version) Then
            Form1.AppendText(JL.Results, True)
            Form1.AppendText("Skipping flash")
            Return True
        End If

        Form1.AppendText("Programming flash with version " + Version, True)
        If (Not JL.Flash(Version)) Then
            Form1.AppendText(JL.Results, True)
            Return False
        End If
        Form1.AppendText(JL.Results, True)

        Form1.AppendText("Re-Verifying Flash", True)
        If Not JL.Verify(Version) Then
            Form1.AppendText(JL.Results, True)
        End If
        Form1.AppendText(JL.Results, True)

        Return True
    End Function


End Class