Partial Class Tests
    Public Shared UUT_LED As Integer
    Public Shared UUT_STATE As Integer
    Public Shared LedTimer As Timer

    Public Shared Function Test_LED() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        ' Dim DR As DialogResult

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        UUT_LED = LED.ALARM
        UUT_STATE = State._ON
        LedTimer = New Timer
        LedTimer.Interval = 100
        AddHandler LedTimer.Tick, AddressOf LedFlasher_Tick
        LedTimer.Start()
        'DR = MessageBox.Show("Is the alarm LED flashing RED", Form1.Serial_Number_Entry.Text + ":  ALARM LED RED?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated the ALARM LED is not flashing RED.")
        '    LedTimer.Stop()
        '    Return False
        'End If
        If Not YesNo("Is the alarm LED flashing RED", "ALARM LED RED?") Then
            Form1.AppendText("Operator indicated the ALARM LED is not flashing RED.")
            LedTimer.Stop()
            Return False
        End If
        LedTimer.Stop()
        T.SetLED(UUT_LED, State._OFF, True)

        UUT_LED = LED.SERVICE
        LedTimer.Start()
        'DR = MessageBox.Show("Is the service LED flashing BLUE", Form1.Serial_Number_Entry.Text + ":  SERVICE LED BLUE?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated the SERVICE LED is not flashing BLUE.")
        '    LedTimer.Stop()
        '    Return False
        'End If
        If Not YesNo("Is the service LED flashing BLUE", "SERVICE LED BLUE?") Then
            Form1.AppendText("Operator indicated the SERVICE LED is not flashing BLUE.")
            LedTimer.Stop()
            Return False
        End If
        LedTimer.Stop()
        T.SetLED(UUT_LED, State._OFF, True)

        UUT_LED = LED.POWER
        LedTimer.Start()
        'DR = MessageBox.Show("Is the power LED flashing GREEN", Form1.Serial_Number_Entry.Text + ":  POWER LED GREEN?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated the POWER LED is not flashing GREEN.")
        '    LedTimer.Stop()
        '    Return False
        'End If
        If Not YesNo("Is the power LED flashing GREEN", "POWER LED GREEN?") Then
            Form1.AppendText("Operator indicated the POWER LED is not flashing GREEN.")
            LedTimer.Stop()
            Return False
        End If
        LedTimer.Stop()
        T.SetLED(UUT_LED, State._OFF, True)

        Return True
    End Function


    Public Shared Sub LedFlasher_Tick(sender As System.Object, e As System.EventArgs)
        Dim T As New TM1

        LedTimer.Enabled = False
        If T.SetLED(UUT_LED, UUT_STATE, True) Then
            LedTimer.Enabled = True
        End If
        If UUT_STATE = State._ON Then
            UUT_STATE = State._OFF
        Else
            UUT_STATE = State._ON
        End If

    End Sub
End Class