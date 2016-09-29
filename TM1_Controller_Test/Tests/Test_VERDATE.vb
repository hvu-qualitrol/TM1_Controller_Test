Partial Class Tests
    Public Shared Function Test_VERDATE() As Boolean
        Dim results As ReturnResults
        Dim T As New TM1
        Dim SF As New SerialFunctions

        ' Power down and then power up the board to make sure that the date/time
        ' is running correctly on the battery. This is to catch any issue with battery
        If Not YesNo("Power down the controller board", "Power down?") Then
            Form1.AppendText("Operator indicated problem powering down.")
            Return False
        End If

        If Not YesNo("Power up the controller board", "Power up?") Then
            Form1.AppendText("Operator indicated problem powering up.")
            Return False
        End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If
        Threading.Thread.Sleep(3000)

        If Not T.VerifyDate() Then
            Return False
        End If
        Return True
    End Function

End Class