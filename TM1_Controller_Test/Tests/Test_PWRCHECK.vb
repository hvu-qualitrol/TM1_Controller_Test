Partial Class Tests
    Public Shared Function Test_PWRCHECK() As Boolean
        'Dim DF As New DAQ_functions
        'Dim Reading As Double
        Dim out3_daq_channel As Integer = 1
        Dim results As ReturnResults
        Dim T As New TM1
        Dim SF As New SerialFunctions
        Dim ADC_Readings As New Hashtable

        'If Not DF.InitializeDAQ Then
        '    Form1.AppendText(DF.GetErrorMsg, True)
        'End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        System.Threading.Thread.Sleep(50)
        If Not T.ReadADC(ADC_Readings) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        For Each k In ADC_Readings.Keys
            Form1.AppendText(k + "=" + ADC_Readings(k).ToString)
        Next

        If Not Specs("24VREF").CompareToSpec(ADC_Readings("24VREF")) Then
            Form1.AppendText("24VREF" + Specs("24VREF").ErrorMsg)
            Return False
        End If

        If Not Specs("5VREF").CompareToSpec(ADC_Readings("5VREF")) Then
            Form1.AppendText("5VREF" + Specs("5VREF").ErrorMsg)
            Return False
        End If

        Return True
    End Function
End Class