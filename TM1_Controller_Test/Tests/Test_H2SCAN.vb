Partial Class Tests
    Public Shared Function Test_H2SCAN() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        ' Do an initial connection to the sensor and then abort to work around issue with running
        ' sensor -d for the first time after booting.
        If Not SF.Cmd(SerialPort, Response, "sensor -d", 10, "Press CTRL-D to break out of this mode.") Then
            Form1.AppendText("Sent cmd 'sensor -d', did not see response 'Press CTRL-D to break out of this mode.'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)
        If Not SF.Cmd(SerialPort, Response, Chr(4), 10, "> ", False, True, False) Then
            Form1.AppendText("Couldn't exit from sensor debug mode, expected to see '> '")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        CommonLib.Delay(2)
        If Not SF.Cmd(SerialPort, Response, "sensor -d", 10, "Press CTRL-D to break out of this mode.") Then
            Form1.AppendText("Sent cmd 'sensor -d', did not see response 'Press CTRL-D to break out of this mode.'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        CommonLib.Delay(5)
        If Not SF.Cmd(SerialPort, Response, Chr(27), 10, "Messages", True, False, False) Then
            Form1.AppendText("Didn't see 'Messages'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        If Not SF.Cmd(SerialPort, Response, Chr(27), 10, "H2scan: ", False, True, False) Then
            Form1.AppendText("Didn't get HS scan prompt 'H2scan: '")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        If Not SF.Cmd(SerialPort, Response, Chr(4), 10, "> ", False, True, False) Then
            Form1.AppendText("Couldn't exit from sensor debug mode, expected to see '> '")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        Return True
    End Function
End Class