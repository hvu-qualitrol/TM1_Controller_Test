Partial Class Tests
    Public Shared Function Test_RTC() As Boolean
        Dim freq_counter As New BK1856D
        Dim startTime As DateTime
        Dim reading_valid As Boolean
        Dim reading As Double
        Dim prev_reading As Double
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim rtc As New rtc
        Dim rtc_info As Hashtable
        Dim capacitance As Integer
        Dim best_capacitance As Integer
        Dim frequency_at_best_capacitance As Double
        Dim Response As String

        Form1.AppendText("Opening comm to freq counter")
        If Not freq_counter.Open(Form1.BK1856D_com_ComboBox.Text) Then
            Form1.AppendText(freq_counter.ErrorMsg)
            Return False
        End If
        Form1.AppendText("Successfully opened comm to freq counter")

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            freq_counter.Close()
            Return False
        End If

        If Not rtc.Clear() Then
            Form1.AppendText(rtc.ErrorMsg)
            freq_counter.Close()
            Return False
        End If

        If Not rtc.GetRTC(rtc_info) Then
            Form1.AppendText(rtc.ErrorMsg)
            freq_counter.Close()
            Return False
        End If

        prev_reading = 666.666
        best_capacitance = 0
        frequency_at_best_capacitance = 666.666
        CommonLib.Delay(5)
        For capacitance = 0 To 30 Step 2
            If Not rtc.SetCapacitance(capacitance) Then
                Form1.AppendText(rtc.ErrorMsg)
                freq_counter.Close()
                Return False
            End If
            'CommonLib.Delay(2)
            startTime = Now
            reading_valid = False
            While (Not reading_valid And Now.Subtract(startTime).TotalSeconds < 30)
                System.Threading.Thread.Sleep(200)
                If Not freq_counter.Read(reading) Then
                    Form1.AppendText(freq_counter.ErrorMsg)
                    freq_counter.Close()
                    Return False
                End If

                If reading > 0.9 And reading < 1.1 And reading < prev_reading Then
                    reading_valid = True
                End If
            End While
            Form1.AppendText("Capacitance = " + capacitance.ToString + " RTC frequency = " + reading.ToString + " Hz")
            If Not reading_valid Then
                Form1.AppendText("Timeout waiting for the reading to be valid ( > 0.9, < 1.1 and change from the prev capacitance value")
                freq_counter.Close()
                Return False
            End If
            If Math.Abs(1 - reading) < Math.Abs(1 - frequency_at_best_capacitance) Then
                frequency_at_best_capacitance = reading
                best_capacitance = capacitance
            End If
            prev_reading = reading
        Next

        Form1.AppendText("Best capacitance = " + best_capacitance.ToString + ", Frequency = " + frequency_at_best_capacitance.ToString)
        Form1.AppendText("Setting capacitance to " + best_capacitance.ToString + " and verifying frequency")
        If Not rtc.SetCapacitance(best_capacitance) Then
            Form1.AppendText(rtc.ErrorMsg)
            freq_counter.Close()
            Return False
        End If
        startTime = Now
        reading_valid = False
        While (Not reading_valid And Now.Subtract(startTime).TotalSeconds < 10)
            System.Threading.Thread.Sleep(200)
            If Not freq_counter.Read(reading) Then
                Form1.AppendText(freq_counter.ErrorMsg)
                freq_counter.Close()
                Return False
            End If

            If Math.Abs(reading - frequency_at_best_capacitance) < 0.000002 Then
                reading_valid = True
            End If
        End While
        If Not reading_valid Then
            Form1.AppendText("timeout waiting for frequency to settle to " + frequency_at_best_capacitance.ToString + " +/-.000001")
            freq_counter.Close()
            Return False
        End If

        If Not rtc.SetOffset(reading) Then
            Form1.AppendText(rtc.ErrorMsg)
            freq_counter.Close()
            Return False
        End If
        'Form1.AppendText("Opening comm to freq counter")
        'If Not freq_counter.Open(Form1.BK1856D_com_ComboBox.Text) Then
        '    Form1.AppendText(freq_counter.ErrorMsg)
        '    Return False
        'End If
        'Form1.AppendText("Successfully opened comm to freq counter")


        'For i = 0 To 100
        '    Application.DoEvents()
        '    startTime = Now
        '    reading_valid = False
        '    While (Not reading_valid And Now.Subtract(startTime).TotalSeconds < 30)
        '        If Not freq_counter.Read(reading) Then
        '            Form1.AppendText(freq_counter.ErrorMsg)
        '            freq_counter.Close()
        '            Return False
        '        End If

        '        If reading > 0.9 And reading < 1.1 Then
        '            reading_valid = True
        '        End If
        '    End While
        '    Form1.AppendText("RTC frequency = " + reading.ToString + " Hz")
        'Next
        'startTime = Now
        'reading_valid = False
        'While (Not reading_valid And Now.Subtract(startTime).TotalSeconds < 30)
        '    If Not freq_counter.Read(reading) Then
        '        Form1.AppendText(freq_counter.ErrorMsg)
        '        Return False
        '    End If

        '    If reading > 0.9 And reading < 1.1 Then
        '        reading_valid = True
        '    End If
        'End While
        'Form1.AppendText("RTC frequency = " + reading.ToString + " Hz")
        If Not freq_counter.Close() Then
            Form1.AppendText(freq_counter.ErrorMsg)
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "rtc", 10) Then
            Form1.AppendText("Problem sending cmd 'rtc'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        If Not SF.Cmd(SerialPort, Response, "cat config.dat", 10) Then
            Form1.AppendText("Problem sending cmd 'cat config.dat'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)

        Return reading_valid
    End Function
End Class
