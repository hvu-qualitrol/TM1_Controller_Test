Partial Class Tests
    Public Shared Function Test_PUMP() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim PumpData As Hashtable
        Dim T As New TM1
        Dim PumpTest As Hashtable
        Dim PumpTestSequence As New ArrayList

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "STOP")
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 0.0)
        PumpTest.Add("UL", 0.0)
        PumpTestSequence.Add(PumpTest)

        'PumpTest = New Hashtable
        'PumpTest.Add("DIR", "FORWARD")
        'PumpTest.Add("PERCENTAGE", 20)
        'PumpTest.Add("RPM", 0.0)
        'PumpTest.Add("LL", 2.5)
        'PumpTest.Add("UL", 2.7)
        'PumpTestSequence.Add(PumpTest)

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "FORWARD")
        PumpTest.Add("PERCENTAGE", 40)
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 4.9)
        PumpTest.Add("UL", 5.4)
        PumpTestSequence.Add(PumpTest)

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "FORWARD")
        PumpTest.Add("PERCENTAGE", 60)
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 6.7)
        PumpTest.Add("UL", 7.1)
        PumpTestSequence.Add(PumpTest)

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "STOP")
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 0.0)
        PumpTest.Add("UL", 0.0)
        PumpTestSequence.Add(PumpTest)

        'PumpTest = New Hashtable
        'PumpTest.Add("DIR", "REVERSE")
        'PumpTest.Add("PERCENTAGE", 20)
        'PumpTest.Add("RPM", 0.0)
        'PumpTest.Add("LL", 2.5)
        'PumpTest.Add("UL", 2.7)
        'PumpTestSequence.Add(PumpTest)

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "REVERSE")
        PumpTest.Add("PERCENTAGE", 40)
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 4.9)
        PumpTest.Add("UL", 5.4)
        PumpTestSequence.Add(PumpTest)

        PumpTest = New Hashtable
        PumpTest.Add("DIR", "REVERSE")
        PumpTest.Add("PERCENTAGE", 60)
        PumpTest.Add("RPM", 0.0)
        PumpTest.Add("LL", 6.7)
        PumpTest.Add("UL", 7.1)
        PumpTestSequence.Add(PumpTest)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        For Each PumpTest In PumpTestSequence
            If PumpTest("DIR") = "STOP" Then
                Form1.AppendText("Stopping pump")
                If Not T.PumpOff() Then
                    Form1.AppendText(T.ErrorMsg)
                    Return False
                End If
                System.Threading.Thread.Sleep(50)
            Else
                Form1.AppendText("Driving pump " + PumpTest("DIR") + " at speed " + PumpTest("PERCENTAGE").ToString)
                If Not T.SetPumpSpeed(PumpTest("DIR"), PumpTest("PERCENTAGE")) Then
                    Form1.AppendText(T.ErrorMsg)
                    Return False
                End If
                CommonLib.Delay(5)
            End If

            PumpData = T.GetPumpRpm()
            If Not PumpData("STATUS") Then
                Form1.AppendText("Problem getting pump RPM")
                Form1.AppendText(PumpData("ERROR_MESSAGE"))
                Return False
            End If
            Form1.AppendText("PUMP RPM = " + PumpData("RPM").ToString)
            Form1.AppendText("PUMP DIR = " + PumpData("DIR"))
            If PumpTest("DIR") = "STOP" Then
                Continue For
            End If
            If (Not PumpTest("DIR") = "STOP" And
                (Not PumpData("DIR") = PumpTest("DIR")) Or
                (PumpData("RPM") > PumpTest("UL")) Or
                (PumpData("RPM") < PumpTest("LL"))) Then
                Form1.AppendText("Expected DIR=" + PumpTest("DIR") + " and RPM between " + PumpTest("LL").ToString + " and " + PumpTest("UL").ToString)
                Return False
            End If
        Next
        Return True

        Form1.AppendText("Turning pump off")
        If Not T.PumpOff() Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        PumpData = T.GetPumpRpm()
        If Not PumpData("STATUS") Then
            Form1.AppendText("Problem getting pump RPM")
            Return False
        End If
        Form1.AppendText("PUMP RPM = " + PumpData("RPM").ToString)
        Form1.AppendText("PUMP DIR = " + PumpData("DIR"))
        'If (Not PumpData("DIR") = "STOP") Or (Not PumpData("RPM") = 0.0) Then
        '    Form1.AppendText("Expected DIR=STOP and RPM=0")
        '    Return False
        'End If

        If Not T.SetPumpSpeed("FOR", 20) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        CommonLib.Delay(3)
        For i = 0 To 100
            PumpData = T.GetPumpRpm()
            If Not PumpData("STATUS") Then
                Form1.AppendText("Problem getting pump RPM")
                Return False
            End If
            Form1.AppendText("PUMP RPM = " + PumpData("RPM").ToString)
            Form1.AppendText("PUMP DIR = " + PumpData("DIR"))
        Next
        'PumpData = T.GetPumpRpm()
        'If Not PumpData("STATUS") Then
        '    Form1.AppendText("Problem getting pump RPM")
        '    Return False
        'End If
        'Form1.AppendText("PUMP RPM = " + PumpData("RPM").ToString)
        'Form1.AppendText("PUMP DIR = " + PumpData("DIR"))
        If (Not PumpData("DIR") = "FOR") Or (Not PumpData("RPM") = 0.0) Then
            Form1.AppendText("Expected DIR=FOR and RPM=0")
            Return False
        End If

        Return True


    End Function
End Class