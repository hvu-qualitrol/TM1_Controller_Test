Partial Class Tests
    Public Shared Function Test_H2SCAN_SUB_ASSEMBLY() As Boolean
        Dim SensorInfo As Hashtable
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim SerialMode As H2SCAN_debug.H2SCAN_SI_MODE
        Dim T As New TM1
        Dim DB As New DB
        Dim NewLinkInfo As Hashtable

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not (h2scan.GetSerialMode(SerialMode)) Then
            Form1.AppendText(h2scan.ErrorMsg)
            Return False
        End If
        Form1.AppendText("H2SCAN Serial Mode = " + SerialMode.ToString)

        If Not SerialMode = H2SCAN_debug.H2SCAN_SI_MODE.MODBUS Then
            If Not h2scan.Open() Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If

            If Not h2scan.VerifyModbusID() Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If

            If Not h2scan.SetSerialMode(H2SCAN_debug.H2SCAN_SI_MODE.MODBUS) Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If
            CommonLib.Delay(5)

            results = SF.Reboot(SerialPort)
            If Not results.PassFail Then
                Form1.AppendText("Problem rebooting")
                Form1.AppendText(SF.ErrorMsg)
                Return False
            End If

            If Not (h2scan.GetSerialMode(SerialMode)) Then
                Form1.AppendText(h2scan.ErrorMsg)
                Return False
            End If
            Form1.AppendText("H2SCAN Serial Mode = " + SerialMode.ToString)

            If Not SerialMode = H2SCAN_debug.H2SCAN_SI_MODE.MODBUS Then
                Form1.AppendText("Problem setting H2SCAN serial mode to MODBUS")
                Return False
            End If
        End If




        'If Not h2scan.Close() Then
        '    Form1.AppendText(h2scan.ErrorMsg)
        '    Return False
        'End If


        If Not T.GetSensors(SensorInfo) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        For Each k In SensorInfo.Keys
            Form1.AppendText(k + " = " + SensorInfo(k))
        Next
        If Not SensorInfo.Contains("Sensor Model") Then
            Form1.AppendText("missing 'Sensor Model' from sensor ouput")
            Return False
        End If
        If Not SensorInfo("Sensor Model") = "4400-BB01" And
            Not SensorInfo("Sensor Model") = "4400-BB01-P" And
            Not SensorInfo("Sensor Model") = "4400-BB02-P1" And
            Not SensorInfo("Sensor Model") = "4400-BBO2-P1" And
            Not SensorInfo("Sensor Model") = "4400-BB02" Then
            Form1.AppendText("Expected 'Sensor Model' = 4400-BB01, 4400-BB01-P, or 4400-BB02-P1, 4400-BBO2-P1, or 4400-BB02")
            Return False
        End If

        NewLinkInfo = New Hashtable
        NewLinkInfo.Add("TM1", TM1_SN)
        NewLinkInfo.Add("H2SCAN", SensorInfo("Sensor Product Serial"))
        If Not Form1.ListBox_TestMode.Text = "DEBUG" Then
            If Not DB.AddLinkData(NewLinkInfo) Then
                Form1.AppendText("Problem updating TM1_Link_Table")
                Form1.AppendText(DB.ErrorMsg)
                Return False
            End If
        End If

        Return True
    End Function
End Class