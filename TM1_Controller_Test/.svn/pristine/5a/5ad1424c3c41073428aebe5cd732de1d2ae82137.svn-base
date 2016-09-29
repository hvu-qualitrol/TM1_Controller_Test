Partial Class Tests
    Public Shared Function Test_OILBLOCK_HEATER() As Boolean
        'Dim DF As New DAQ_functions
        Dim RDC As New RemoteDaqClient
        Dim Reading As Double
        Dim T As New TM1
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim oilblock_daq_chan As Integer = 3
        Dim analogbd_daq_chan As Integer = 2
        Dim HeaterResultsStarting As Hashtable
        Dim HeaterResultsEnding As Hashtable

        If Not RDC.LockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
            Return False
        End If
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

        CommonLib.Delay(12)
        Form1.AppendText("Driving PWM for 120 seconds")
        If Not T.SetPWM("ANALOGBD", 0) Then
            T.SetPWM_AUTO()
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If Not T.SetPWM("OILBLOCK", 0) Then
            T.SetPWM_AUTO()
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If Not T.Heater(HeaterResultsStarting) Then
            T.SetPWM_AUTO()
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("AnalogBd starting temp = " + HeaterResultsStarting("AnalogBdTemp").ToString)
        Form1.AppendText("OilBlockTemp starting temp = " + HeaterResultsStarting("OilBlockTemp").ToString)
        CommonLib.Delay(120)
        If Not T.Heater(HeaterResultsEnding) Then
            T.SetPWM_AUTO()
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("AnalogBd ending temp = " + HeaterResultsEnding("AnalogBdTemp").ToString)
        Form1.AppendText("OilBlockTemp ending temp = " + HeaterResultsEnding("OilBlockTemp").ToString)

        If Not T.SetPWM_AUTO() Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If HeaterResultsStarting("AnalogBdTemp") - HeaterResultsEnding("AnalogBdTemp") < 1 Then
            Form1.AppendText("Expected AnalogBdTemp rise > 1C")
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If HeaterResultsStarting("OilBlockTemp") - HeaterResultsEnding("OilBlockTemp") < 1 Then
            Form1.AppendText("Expected OilBlockTemp rise > 1C")
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("OILBLOCK", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        'DF.GetAnalogReading(oilblock_daq_chan, Reading)
        If Not RDC.GetAnalogReading(oilblock_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("OILBLOCK TEC DC Voltage = " + Reading.ToString)
        If Reading < -2.1 Or Reading > -0.5 Then
            Form1.AppendText("Expected between -2.1 and -0.5")
            T.SetTEC("OILBLOCK", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "HEAT", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("OILBLOCK", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        If Not RDC.GetAnalogReading(oilblock_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            T.SetTEC("OILBLOCK", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        'DF.GetAnalogReading(oilblock_daq_chan, Reading)
        Form1.AppendText("OILBLOCK TEC DC Voltage = " + Reading.ToString)
        If Reading < 0.5 Or Reading > 2.5 Then
            Form1.AppendText("Expected between 0.5 and 2.5")
            T.SetTEC("OILBLOCK", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            T.SetTEC_AUTO()
            Return False
        End If

        If Not T.SetTEC_AUTO() Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        'CommonLib.Delay(120)
        'DF.GetAnalogReading(analogbd_daq_chan, Reading)
        If Not RDC.GetAnalogReading(analogbd_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("ANALOGBD TEC DC Voltage = " + Reading.ToString)
        If Reading < -2.1 Or Reading > -0.1 Then
            Form1.AppendText("Expected between -2.1 and -0.1")
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "HEAT", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        'CommonLib.Delay(120)
        'DF.GetAnalogReading(analogbd_daq_chan, Reading)
        If Not RDC.GetAnalogReading(analogbd_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("ANALOGBD TEC DC Voltage = " + Reading.ToString)
        If Reading < 0.1 Or Reading > 2.5 Then
            Form1.AppendText("Expected between 0.1 and 2.1")
            T.SetTEC("ANALOGBD", "COOL", 0)
            T.SetTEC_AUTO()
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC_AUTO() Then
            Form1.AppendText(T.ErrorMsg)
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

    Public Shared Function Test_OILBLOCK_HEATER0() As Boolean
        'Dim DF As New DAQ_functions
        Dim RDC As New RemoteDaqClient
        Dim Reading As Double
        Dim T As New TM1
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        'Dim TecStatus As Hashtable
        Dim oilblock_daq_chan As Integer = 3
        Dim analogbd_daq_chan As Integer = 2
        Dim HeaterResultsStarting As Hashtable
        Dim HeaterResultsEnding As Hashtable

        If Not RDC.LockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
            Return False
        End If
        If Not RDC.InitializeDAQ Then
            Form1.AppendText(RDC.ErrorMsg, True)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        'If Not DF.InitializeDAQ Then
        '    Form1.AppendText(DF.GetErrorMsg, True)
        '    Return False
        'End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        CommonLib.Delay(12)
        Form1.AppendText("Driving PWM for 40 seconds")
        If Not T.Heater(HeaterResultsStarting) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If Not T.Heater(HeaterResultsStarting) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("AnalogBd starting temp = " + HeaterResultsStarting("AnalogBdTemp").ToString)
        Form1.AppendText("OilBlockTemp starting temp = " + HeaterResultsStarting("OilBlockTemp").ToString)
        If Not T.SetPWM("ANALOGBD", 100) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If Not T.SetPWM("OILBLOCK", 100) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(80)
        If Not T.Heater(HeaterResultsEnding) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("AnalogBd ending temp = " + HeaterResultsEnding("AnalogBdTemp").ToString)
        Form1.AppendText("OilBlockTemp ending temp = " + HeaterResultsEnding("OilBlockTemp").ToString)
        If Not T.SetPWM("ANALOGBD", 0) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If Not T.SetPWM("OILBLOCK", 0) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        If HeaterResultsEnding("AnalogBdTemp") - HeaterResultsStarting("AnalogBdTemp") < 3 Then
            Form1.AppendText("Expected AnalogBdTemp rise > 3C")
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If HeaterResultsEnding("OilBlockTemp") - HeaterResultsStarting("OilBlockTemp") < 3 Then
            Form1.AppendText("Expected OilBlockTemp rise > 3C")
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("OILBLOCK", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        'DF.GetAnalogReading(oilblock_daq_chan, Reading)
        If Not RDC.GetAnalogReading(oilblock_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("OILBLOCK TEC DC Voltage = " + Reading.ToString)
        If Reading < -2.1 Or Reading > -1.3 Then
            Form1.AppendText("Expected between -2.1 and -1.3")
            T.SetTEC("OILBLOCK", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "HEAT", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("OILBLOCK", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        CommonLib.Delay(5)
        If Not RDC.GetAnalogReading(oilblock_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        'DF.GetAnalogReading(oilblock_daq_chan, Reading)
        Form1.AppendText("OILBLOCK TEC DC Voltage = " + Reading.ToString)
        If Reading < 1.3 Or Reading > 2.5 Then
            Form1.AppendText("Expected between 1.3 and 2.1")
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        'CommonLib.Delay(5)
        CommonLib.Delay(50)
        'DF.GetAnalogReading(analogbd_daq_chan, Reading)
        If Not RDC.GetAnalogReading(analogbd_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("ANALOGBD TEC DC Voltage = " + Reading.ToString)
        'If Reading < -2.1 Or Reading > -1.5 Then
        If Reading < -2.1 Or Reading > -0.2 Then
            'Form1.AppendText("Expected between -2.1 and -1.3")
            Form1.AppendText("Expected between -2.1 and -0.20")
            T.SetTEC("ANALOGBD", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "HEAT", 100) Then
            Form1.AppendText(T.ErrorMsg)
            T.SetTEC("ANALOGBD", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        'CommonLib.Delay(5)
        CommonLib.Delay(50)
        RDC.GetAnalogReading(analogbd_daq_chan, Reading)
        If Not RDC.GetAnalogReading(analogbd_daq_chan, Reading) Then
            Form1.AppendText(RDC.ErrorMsg)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If
        Form1.AppendText("ANALOGBD TEC DC Voltage = " + Reading.ToString)
        'If Reading < 1.5 Or Reading > 2.5 Then
        If Reading < 0.2 Or Reading > 2.5 Then
            'Form1.AppendText("Expected between 1.3 and 2.1")
            Form1.AppendText("Expected between 0.2 and 2.5")
            T.SetTEC("ANALOGBD", "COOL", 0)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        If Not T.SetTEC("ANALOGBD", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
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