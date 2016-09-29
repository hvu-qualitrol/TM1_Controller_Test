Imports System.IO.Ports

Partial Class Tests
    'Shared BG As Background

    Public Shared Function Test_THERMALS_LOW() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim Parent As Control = Form1.GetControl
        'Dim BG As New Background(Parent, AddressOf UpdateHeaterTable)
        Dim Response As String
        'Dim HeaterThread As New System.Threading.Thread(AddressOf BG.LogHeater)
        'Dim row_cnt As Integer
        Dim StartTime As DateTime
        Dim conditions As Hashtable
        Dim cond As condition
        Dim NewConfig As New Hashtable
        Dim RebootNeeded As Boolean
        Dim RowDateTime As DateTime
        Dim WhereClause As String
        Dim ts_60s As New TimeSpan(0, 0, 60)
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim TEC_COOL_analogbd_temp, TEC_COOL_oilblock_temp, TEC_COOL_analogbd_pwm, TEC_COOL_oilblock_pwm, TEC_COOL_analogbd_stdev, TEC_COOL_oilblock_stdev As Double
        Dim SpecFailed As Boolean

        BG = New Background(Parent, AddressOf UpdateHeaterTable)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        NewConfig.Add("analogbd.setpoint", "20.000")
        NewConfig.Add("oilblock.setpoint", "20.000")
        NewConfig.Add("CLI_OVER_TMCOM1.ENABLE", "TRUE")
        NewConfig.Add("TMCOM1.PROTOCOL", "CLI")
        NewConfig.Add("TMCOM1.MODE", "RS232")
        If Not T.SetConfig(NewConfig, RebootNeeded, True) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        CommonLib.Delay(5)
        If Not SF.Cmd(SerialPort, Response, "sensor off", 5) Then
            Form1.AppendText("Problem sending command 'sensor off'")
            Return False
        End If

        BG.SerialPortName = Form1.Tmcom1ComboBox.Text
        BG.NewHeaterTable()
        Form1.DataGridView1.DataSource = BG.HeaterDT
        For Each col In Form1.DataGridView1.Columns
            If Not (col.Name = "Timestamp" Or col.Name = "OilBlockTemp" Or
                    col.Name = "OilBlockPWM" Or col.Name = "AnalogBdTemp" Or
                    col.Name = "AnalogBdPWM") Then
                Form1.DataGridView1.Columns(col.Name).Visible = False
            Else
                Form1.DataGridView1.Columns(col.Name).Visible = True
                Form1.DataGridView1.Columns(col.Name).Width = 50
            End If
        Next

        CommonLib.Delay(10)
        If Not BG.StartLogging() Then
            Return False
        End If
        'If Not SF.Cmd(SerialPort, Response, "sensor off", 5) Then
        '    Form1.AppendText("Problem sending command 'sensor off'")
        '    Return False
        'End If

        conditions = New Hashtable

        cond = New condition
        cond.expected = 25.0
        conditions.Add("AnalogBdTemp", cond)

        cond = New condition
        cond.expected = 25.0
        conditions.Add("OilBlockTemp", cond)

        Form1.AppendText("Waiting for AnalogBdTemp & OilBlockTmp to reach 25C...")
        If Not BG.GetZonesToTemp(conditions, 1000, "TEC_COOL") Then
            Form1.AppendText("Zones did not reach initial target temperatures")
            BG.EndLogging()
            csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\TEC_COOL." + TimeStamp + ".csv"
            If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
                Form1.AppendText("Problem creating csv file " + csv_filepath)
            End If
            Return False
        End If

        conditions = New Hashtable

        cond = New condition
        cond.expected = 20
        cond.limit = 0.2
        cond.stop_when_stable = True
        'cond.expected = 22.0
        'cond.limit = 0.2
        conditions.Add("AnalogBdTemp", cond)

        cond = New condition
        cond.pwm = 0.5
        conditions.Add("AnalogBdPWM", cond)

        cond = New condition
        cond.expected = 20
        cond.limit = 0.2
        cond.stop_when_stable = True
        'cond.expected = 21.0
        'cond.limit = 0.2
        conditions.Add("OilBlockTemp", cond)

        cond = New condition
        cond.pwm = 0.5
        conditions.Add("OilBlockPWM", cond)

        If Not T.SetTEC_AUTO Then
            Form1.AppendText(T.ErrorMsg)
            BG.EndLogging()
            Return False
        End If
        If Not T.SetPWM_AUTO Then
            Form1.AppendText(T.ErrorMsg)
            BG.EndLogging()
            Return False
        End If

        Form1.AppendText("Waiting for AnalogBdTemp to stabilize within 0.2C...")
        Form1.AppendText("Waiting for OilBlockTemp to stabilize within 0.2C...")
        StartTime = Now
        If Not BG.WaitHeaterCondition(conditions, 1400) Then
            Form1.AppendText("Timeout waiting for temperature to stabilize")
            BG.EndLogging()
            csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\TEC_COOL." + TimeStamp + ".csv"
            If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
                Form1.AppendText("Problem creating csv file " + csv_filepath)
            End If
            Return False
        End If
        Form1.AppendText("Stabilization time:  " + Now.Subtract(StartTime).TotalSeconds.ToString)

        Form1.AppendText("CAPTURING 60s OF HEATER DATA")
        StartTime = Now
        While BG.HeaterThread.IsAlive And Now.Subtract(StartTime).TotalSeconds < 60
            Application.DoEvents()
        End While
        BG.EndLogging()

        RowDateTime = BG.HeaterDT.Rows(BG.HeaterDT.Rows.Count - 1)("Timestamp")
        WhereClause = "Timestamp > '" + RowDateTime.Subtract(ts_60s).ToString + "'"

        Form1.AppendText("row count = " + BG.HeaterDT.Rows.Count.ToString + vbCr)

        SpecFailed = False
        TEC_COOL_oilblock_pwm = BG.HeaterDT.Compute("AVG(OilBLockPWM)", WhereClause)
        Form1.AppendText("ave OilBlockPWM = " + TEC_COOL_oilblock_pwm.ToString)
        If Not Specs("TEC_COOL_oilblock_pwm").CompareToSpec(TEC_COOL_oilblock_pwm) Then
            Form1.AppendText(Specs("TEC_COOL_oilblock_pwm").ErrorMsg)
            SpecFailed = True
        End If

        TEC_COOL_analogbd_pwm = BG.HeaterDT.Compute("AVG(AnalogBdPWM)", WhereClause)
        Form1.AppendText("ave AnalogBdPWM = " + TEC_COOL_analogbd_pwm.ToString)
        If Not Specs("TEC_COOL_analogbd_pwm").CompareToSpec(TEC_COOL_analogbd_pwm) Then
            Form1.AppendText(Specs("TEC_COOL_analogbd_pwm").ErrorMsg)
            SpecFailed = True
        End If

        TEC_COOL_oilblock_temp = BG.HeaterDT.Compute("AVG(OilBlockTemp)", WhereClause)
        Form1.AppendText("ave OilBlockTemp = " + TEC_COOL_oilblock_temp.ToString)
        If Not Specs("TEC_COOL_oilblock_temp").CompareToSpec(TEC_COOL_oilblock_temp) Then
            Form1.AppendText(Specs("TEC_COOL_oilblock_temp").ErrorMsg)
            SpecFailed = True
        End If

        TEC_COOL_analogbd_temp = BG.HeaterDT.Compute("AVG(AnalogBdTemp)", WhereClause)
        Form1.AppendText("ave AnalogBdTemp = " + TEC_COOL_analogbd_temp.ToString)
        If Not Specs("TEC_COOL_analogbd_temp").CompareToSpec(TEC_COOL_analogbd_temp) Then
            Form1.AppendText(Specs("TEC_COOL_analogbd_temp").ErrorMsg)
            SpecFailed = True
        End If

        TEC_COOL_oilblock_stdev = BG.HeaterDT.Compute("STDEV(OilBlockTemp)", WhereClause)
        Form1.AppendText("StDev OilBlockTemp = " + TEC_COOL_oilblock_stdev.ToString)
        If Not Specs("TEC_COOL_oilblock_stdev").CompareToSpec(TEC_COOL_oilblock_stdev) Then
            Form1.AppendText(Specs("TEC_COOL_oilblock_stdev").ErrorMsg)
            SpecFailed = True
        End If

        TEC_COOL_analogbd_stdev = BG.HeaterDT.Compute("STDEV(AnalogBdTemp)", WhereClause)
        Form1.AppendText("StDev AnalogBdTemp = " + TEC_COOL_analogbd_stdev.ToString)
        If Not Specs("TEC_COOL_analogbd_stdev").CompareToSpec(TEC_COOL_analogbd_stdev) Then
            Form1.AppendText(Specs("TEC_COOL_analogbd_stdev").ErrorMsg)
            SpecFailed = True
        End If

        Form1.AppendText("RetVal = " + BG.RetVal.ToString + vbCr)

        If Not SF.Cmd(SerialPort, Response, "sensor on", 5) Then
            Form1.AppendText("Problem sending command 'sensor off'")
            Return False
        End If

        If BG.RetVal = False Then
            Form1.AppendText(BG.ErrorMsg)
            Return False
        End If

        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\TEC_COOL." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
        End If

        If SpecFailed Then Return False

        Return True
    End Function

    'Shared Sub UpdateHeaterTable(ByVal status As String)
    '    Form1.DataGridView1.FirstDisplayedCell = Form1.DataGridView1.Rows(Form1.DataGridView1.Rows.Count - 1).Cells(0)
    '    Form1.DataGridView1.Refresh()
    'End Sub
End Class