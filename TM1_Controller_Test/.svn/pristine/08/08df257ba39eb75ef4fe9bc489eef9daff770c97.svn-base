Partial Class Tests
    Public Shared Function Test_H2SCAN_LAB_MODE() As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim op_mode As H2SCAN_debug.H2SCAN_OP_MODE
        Dim Response As String = ""
        Dim Command As String
        Dim pass As Boolean = True

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        ' Get the current config.sensor.mode setting.
        ' If it is configured for LAB mode then check the actual setting
        Form1.AppendText("Check the current config.sensor.mode setting")
        Command = "config get sensor.mode"
        SF.Cmd(SerialPort, Response, Command, 10)
        If Response.ToUpper.Contains("LAB") Then
            Form1.AppendText("Check the current sensor op mode setting")
            h2scan.GetOperatingMode(op_mode)
            If op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
                pass = True
            End If
        Else
            Form1.AppendText("Change config.sensor.mode to lab. Wait 90s for the change takes effect")
            Command = "config -s set sensor.mode lab"
            SF.Cmd(SerialPort, Response, Command, 10)
            CommonLib.Delay(70)
            Form1.AppendText("Check the current sensor op mode setting")

            For i As Integer = 1 To 3
                h2scan.GetOperatingMode(op_mode)
                If op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
                    pass = True
                    Exit For
                Else
                    pass = False
                    If i < 3 Then CommonLib.Delay(30)
                End If
            Next
        End If
        Form1.AppendText("H2SCAN Operating Mode = " + op_mode.ToString)

        ' This is to force H2Scan to change to LAB mode and correctly reset the cycle time
        If Not h2scan.SetOperatingMode(H2SCAN_debug.H2SCAN_OP_MODE.LAB) Then
            pass = False
            Form1.AppendText(h2scan.ErrorMsg)
            Form1.AppendText("h2scan.SetOperatingMode(" + op_mode.ToString + ") failed")
        End If

        Return pass

    End Function

    Public Shared Function Test_H2SCAN_LAB_MODE0() As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim op_mode As H2SCAN_debug.H2SCAN_OP_MODE
        Dim startTime As DateTime


        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not (h2scan.GetOperatingMode(op_mode)) Then
            Form1.AppendText(h2scan.ErrorMsg)
            Return False
        End If
        Form1.AppendText("H2SCAN Operating Mode = " + op_mode.ToString)

        If Not op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
            If Not h2scan.Open() Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If

            If Not h2scan.SetOperatingMode(H2SCAN_debug.H2SCAN_OP_MODE.LAB) Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If

            startTime = Now
            While Not op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB And Now.Subtract(startTime).TotalSeconds < 30
                CommonLib.Delay(5)
                If Not (h2scan.GetOperatingMode(op_mode)) Then
                    Form1.AppendText(h2scan.ErrorMsg)
                    Return False
                End If
                Form1.AppendText("H2SCAN Operating Mode = " + op_mode.ToString)
            End While
            If Not op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
                Form1.AppendText("mode did not change to lab mode")
                Return False
            End If
        End If

        Return True
    End Function
End Class

