Partial Class Tests
    Public Shared Function Test_420() As Boolean
        ' Dim DF As New DAQ_functions
        Dim RDC As New RemoteDaqClient
        Dim T As New TM1
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim ADC_Readings As Hashtable
        Dim Test_Settings() As String = {"0", "5", "10", "20"}
        Dim Command As String
        Dim SpecName As String
        Dim aux As Integer
        Dim aux_port As String
        Dim reading As Double

        If Not RDC.LockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        'If Not DF.InitializeDAQ Then
        '    Form1.AppendText(DF.GetErrorMsg, True)
        'End If
        If Not RDC.InitializeDAQ Then
            Form1.AppendText(RDC.ErrorMsg, True)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If


        For Each aux In {1, 2}
            Form1.AppendText(vbCr + "-----------------------------------")
            aux_port = "AUX" + aux.ToString
            Form1.AppendText("Testing 420 " + aux_port)
            For i = 0 To UBound(Test_Settings)
                Form1.AppendText("Driving 420_OUT" + aux.ToString + " to " + Test_Settings(i) + "mA")
                System.Threading.Thread.Sleep(50)
                Command = "dac " + (aux - 1).ToString + " " + Test_Settings(i) + ".0"
                If Not SF.Cmd(SerialPort, Response, Command, 10) Then
                    Form1.AppendText("Problem sending command " + Command)
                    Form1.AppendText(SF.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
                CommonLib.Delay(5)
                ADC_Readings = New Hashtable
                If Not T.ReadADC(ADC_Readings) Then
                    Form1.AppendText(T.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
                If Not T.ReadADC(ADC_Readings) Then
                    Form1.AppendText(T.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
                Form1.AppendText("4-20mA_IN" + aux.ToString + " = " + ADC_Readings(aux_port).ToString)
                SpecName = "AUX" + aux.ToString + "_" + Test_Settings(i) + "mA"
                If Not Specs(SpecName).CompareToSpec(ADC_Readings(aux_port)) Then
                    Form1.AppendText(aux_port + Specs(SpecName).ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
            Next

            Command = "dac " + (aux - 1).ToString + " 0.0"
            If Not SF.Cmd(SerialPort, Response, Command, 10) Then
                Form1.AppendText("Problem sending command " + Command)
                If Not RDC.UnlockDaq Then
                    Form1.AppendText(RDC.ErrorMsg)
                End If
                Return False
            End If
        Next

        For i = 0 To UBound(Test_Settings)
            Form1.AppendText("Driving 420 Output 3 to " + Test_Settings(i) + ".0")
            Command = "dac 2 " + Test_Settings(i) + ".0"
            SpecName = "420_" + Test_Settings(i) + "mA_250ohm"
            If Not SF.Cmd(SerialPort, Response, Command, 10) Then
                Form1.AppendText("Problem sending command " + Command)
                If Not RDC.UnlockDaq Then
                    Form1.AppendText(RDC.ErrorMsg)
                End If
                Return False
            End If
            System.Threading.Thread.Sleep(250)
            'DF.GetAnalogReading(1, reading)
            If Not RDC.GetAnalogReading(1, reading) Then
                Form1.AppendText("Problem reading DAQ")
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Form1.AppendText("Voltage acrossed 250 ohm resistor = " + reading.ToString)
            If Not Specs(SpecName).CompareToSpec(reading) Then
                Form1.AppendText("420 output 3 " + Specs(SpecName).ErrorMsg)
                If Not RDC.UnlockDaq Then
                    Form1.AppendText(RDC.ErrorMsg)
                End If
                Return False
            End If
        Next
        Command = "dac 2 0.0"
        If Not SF.Cmd(SerialPort, Response, Command, 10) Then
            Form1.AppendText("Problem sending command " + Command)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not RDC.UnlockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
        End If

        Return True
    End Function
End Class