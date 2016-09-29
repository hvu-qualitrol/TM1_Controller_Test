Partial Class Tests
    Public Shared Function Test_RESET() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String
        Dim ResetSeen As Boolean
        ' Dim DR As DialogResult

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        'Find out how many external reset events has occured
        Dim ExtResetEventCount As Integer = 0
        If Not SF.Cmd(SerialPort, Response, "event 2", 5) Then
            Form1.AppendText("failed sending cmd 'event 2'", True)
            Return False
        End If
        For Each Line In Response.Split(Chr(13))
            Form1.AppendText(Line)
            If Line.Contains("SYSTEM RESTART: External Pin") Then
                ExtResetEventCount += 1
            End If
        Next

        If Not YesNo("Press the reset button for 4 seconds to reset board", "PRESS RESET BUTTON") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
            Return False
        End If
        CommonLib.Delay(4)

        'Reconnect to the unit under test
        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        'Find out if the number of the ext reset event has increase by 1
        Dim NewExtResetEventCount As Integer = 0
        If Not SF.Cmd(SerialPort, Response, "event 2", 5) Then
            Form1.AppendText("failed sending cmd 'event 2'", True)
            Return False
        End If
        ResetSeen = False
        For Each Line In Response.Split(Chr(13))
            Form1.AppendText(Line)
            If Line.Contains("SYSTEM RESTART: External Pin") Then
                NewExtResetEventCount += 1
            End If
        Next
        If NewExtResetEventCount = ExtResetEventCount + 1 Then
            ResetSeen = True
        End If
        Form1.AppendText("ExtResetEventCount = " + ExtResetEventCount.ToString + " NewExtResetEventCount = " + NewExtResetEventCount.ToString)
        If Not ResetSeen Then
            Form1.AppendText("Expected to see event type 2 'SYSTEM RESTART: External Pin'")
            Return False
        End If

        Return True
    End Function

    Public Shared Function Test_RESET0() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String
        Dim ResetSeen As Boolean
        ' Dim DR As DialogResult

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "fr -E", 5) Then
            Form1.AppendText("failed sending cmd 'fr -E'", True)
            Return False
        End If
        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        'DR = MessageBox.Show("Press the reset button for 4 seconds to reset board", Form1.Serial_Number_Entry.Text + ":  PRESS RESET BUTTON", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
        '    Return False
        'End If
        If Not YesNo("Press the reset button for 4 seconds to reset board", "PRESS RESET BUTTON") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
            Return False
        End If
        CommonLib.Delay(4)

        If Not SF.Cmd(SerialPort, Response, "event 2", 5) Then
            Form1.AppendText("failed sending cmd 'event 2'", True)
            Return False
        End If
        ResetSeen = False
        For Each Line In Response.Split(Chr(13))
            Form1.AppendText(Line)
            If Line.Contains("SYSTEM RESTART: External Pin") Then
                ResetSeen = True
            End If
        Next
        If Not ResetSeen Then
            Form1.AppendText("Expected to see event type 2 'SYSTEM RESTART: External Pin'")
            Return False
        End If

        Return True
    End Function
End Class