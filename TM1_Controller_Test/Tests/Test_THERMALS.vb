Imports System.IO.Ports
Imports FTD2XX_NET


Partial Class Tests
    Shared BG As Background

    Public Shared Function TestReworkThermals() As Boolean
        Dim Parent As Control = Form1.GetControl
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim NewConfig As New Hashtable
        Dim RebootNeeded As Boolean
        Dim ts_60s As New TimeSpan(0, 0, 60)
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        NewConfig.Add("analogbd.setpoint", "55.000")
        NewConfig.Add("oilblock.setpoint", "46.000")
        NewConfig.Add("CLI_OVER_TMCOM1.ENABLE", "TRUE")
        NewConfig.Add("TMCOM1.PROTOCOL", "CLI")
        NewConfig.Add("TMCOM1.MODE", "RS232")
        If Not T.SetConfig(NewConfig, RebootNeeded, True) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        System.Threading.Thread.Sleep(2000)
        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        'Cycle comport needed for Windows to detect the port properly
        Dim ft232 As New FT232R
        Dim status = ft232.CycleComport(Form1.Tmcom1ComboBox.Text)
        If Not status Then
            Form1.AppendText("Problem cycleComport TMCOM1 serial port")
            Return False
        End If
        System.Threading.Thread.Sleep(2000)

        BG = New Background(Parent, AddressOf UpdateHeaterTable)
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
        If Not BG.StartLogging() Then
            Return False
        End If

        Dim PassFlag As Boolean = True
        If Not BG.GetReworkZonesToTemp(True, 300) Then
            Form1.AppendText("TestHighThermals failed.")
            PassFlag = False
        Else
            Form1.AppendText("TestHighThermals passed.")
        End If
        BG.EndLogging()
        If PassFlag = False Then
            Return False
        End If

        ' Low thermal test block
        BG = New Background(Parent, AddressOf UpdateHeaterTable)
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
        If Not BG.StartLogging() Then
            Return False
        End If

        If Not BG.GetReworkZonesToTemp(False, 300) Then
            PassFlag = False
            Form1.AppendText("TestLowThermals failed.")
        Else
            Form1.AppendText("TestLowThermals passed.")
        End If
        BG.EndLogging()
        Return PassFlag

    End Function

    Public Shared Function Test_THERMALS() As Boolean
        Dim Parent As Control = Form1.GetControl
        ' Dim BG As New Background(Parent, AddressOf UpdateHeaterTable)
        'Dim row_cnt As Integer
        Dim StartTime As DateTime
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim conditions As Hashtable
        Dim cond As condition
        Dim NewConfig As New Hashtable
        Dim RebootNeeded As Boolean
        Dim RowDateTime As DateTime
        Dim WhereClause As String
        Dim ts_60s As New TimeSpan(0, 0, 60)
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim PWM_HEAT_analogbd_temp, PWM_HEAT_oilblock_temp, PWM_HEAT_analogbd_pwm, PWM_HEAT_oilblock_pwm, PWM_HEAT_analogbd_stdev, PWM_HEAT_oilblock_stdev As Double
        Dim SpecFailed As Boolean
        Dim FT232R_device As New FT232R

        Form1.AppendText("Obtaining FTDI lock")
        If Not FT232R_device.LockFtdi() Then
            Form1.AppendText("Problem getting FTDI lock")
            Form1.AppendText(FT232R_device.failure_message)
            Return False
        End If
        Form1.AppendText("Cycling " + Form1.Tmcom1ComboBox.Text)
        If Not FT232R_device.CycleComport(Form1.Tmcom1ComboBox.Text) Then
            Form1.AppendText("Problem cycling " + Form1.Tmcom1ComboBox.Text)
            FT232R_device.UnlockFtdi()
            Return False
        End If
        FT232R_device.UnlockFtdi()
        CommonLib.Delay(30)

        BG = New Background(Parent, AddressOf UpdateHeaterTable)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        NewConfig.Add("analogbd.setpoint", "55.000")
        NewConfig.Add("oilblock.setpoint", "46.000")
        NewConfig.Add("CLI_OVER_TMCOM1.ENABLE", "TRUE")
        NewConfig.Add("TMCOM1.PROTOCOL", "CLI")
        NewConfig.Add("TMCOM1.MODE", "RS232")
        If Not T.SetConfig(NewConfig, RebootNeeded, True) Then
            Form1.AppendText(T.ErrorMsg)
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
        If Not BG.StartLogging() Then
            Return False
        End If

        conditions = New Hashtable

        cond = New condition
        cond.expected = 34.0
        conditions.Add("AnalogBdTemp", cond)

        cond = New condition
        cond.expected = 33.0
        conditions.Add("OilBlockTemp", cond)

        ' Check for initial conditions
        Form1.AppendText("Checking initial conditions: AnalogBdTemp <= 33C & OilBlockTemp <= 33C...")
        If Not BG.CheckInitConditions(conditions) Then
            Form1.AppendText("Failed to meet the initial conditions.")
            BG.EndLogging()
            Return False
        End If

        conditions("AnalogBdTemp").expected = 49.0
        conditions("OilBlockTemp").expected = 40.0

        Form1.AppendText("Waiting for AnalogBdTemp to reach 49C...")
        Form1.AppendText("Waiting for OilBlockTemp to reach 40C...")
        If Not BG.GetZonesToTemp(conditions, 360, "PWM_HEAT") Then
            BG.EndLogging()
            Form1.AppendText("Timeout waiting for temperature to reach the set points (49C, 40C)")
            csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\PWM_HEAT." + TimeStamp + ".csv"
            If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
                Form1.AppendText("Problem creating csv file " + csv_filepath)
            End If
            Return False
        End If

        If Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
            Form1.AppendText(T.ErrorMsg)
            BG.EndLogging()
            Return False
        End If
        If Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
            Form1.AppendText(T.ErrorMsg)
            BG.EndLogging()
            Return False
        End If
        If Not T.SetPWM_AUTO Then
            Form1.AppendText(T.ErrorMsg)
            BG.EndLogging()
            Return False
        End If

        conditions = New Hashtable

        cond = New condition
        cond.expected = 55.0
        cond.limit = 0.4
        cond.check = ">="
        conditions.Add("AnalogBdTemp", cond)

        cond = New condition
        cond.expected = 46.0
        cond.limit = 0.2
        cond.check = ">="
        conditions.Add("OilBlockTemp", cond)

        Form1.AppendText("Waiting for AnalogBdTemp to stabilize within 55C +/- 0.4C...")
        Form1.AppendText("Waiting for OilBlockTemp to stabilize within 46C +/- 0.2C...")
        StartTime = Now
        If Not BG.WaitHeaterCondition(conditions, 1000) Then
            Form1.AppendText("Timeout waiting for temperature to stabilize")
            BG.EndLogging()
            csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\PWM_HEAT." + TimeStamp + ".csv"
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
        PWM_HEAT_oilblock_pwm = BG.HeaterDT.Compute("AVG(OilBLockPWM)", WhereClause)
        Form1.AppendText("ave OilBlockPWM = " + PWM_HEAT_oilblock_pwm.ToString)
        If Not Specs("PWM_HEAT_oilblock_pwm").CompareToSpec(PWM_HEAT_oilblock_pwm) Then
            Form1.AppendText(Specs("PWM_HEAT_oilblock_pwm").ErrorMsg)
            SpecFailed = True
        End If

        PWM_HEAT_analogbd_pwm = BG.HeaterDT.Compute("AVG(AnalogBdPWM)", WhereClause)
        Form1.AppendText("ave AnalogBdPWM = " + PWM_HEAT_analogbd_pwm.ToString)
        If Not Specs("PWM_HEAT_analogbd_pwm").CompareToSpec(PWM_HEAT_analogbd_pwm) Then
            Form1.AppendText(Specs("PWM_HEAT_analogbd_pwm").ErrorMsg)
            SpecFailed = True
        End If

        PWM_HEAT_oilblock_temp = BG.HeaterDT.Compute("AVG(OilBlockTemp)", WhereClause)
        Form1.AppendText("ave OilBlockTemp = " + PWM_HEAT_oilblock_temp.ToString)
        If Not Specs("PWM_HEAT_oilblock_temp").CompareToSpec(PWM_HEAT_oilblock_temp) Then
            Form1.AppendText(Specs("PWM_HEAT_oilblock_temp").ErrorMsg)
            SpecFailed = True
        End If

        PWM_HEAT_analogbd_temp = BG.HeaterDT.Compute("AVG(AnalogBdTemp)", WhereClause)
        Form1.AppendText("ave AnalogBdTemp = " + PWM_HEAT_analogbd_temp.ToString)
        If Not Specs("PWM_HEAT_analogbd_temp").CompareToSpec(PWM_HEAT_analogbd_temp) Then
            Form1.AppendText(Specs("PWM_HEAT_analogbd_temp").ErrorMsg)
            SpecFailed = True
        End If

        PWM_HEAT_oilblock_stdev = BG.HeaterDT.Compute("STDEV(OilBlockTemp)", WhereClause)
        Form1.AppendText("StDev OilBlockTemp = " + PWM_HEAT_oilblock_stdev.ToString)
        If Not Specs("PWM_HEAT_oilblock_stdev").CompareToSpec(PWM_HEAT_oilblock_stdev) Then
            Form1.AppendText(Specs("PWM_HEAT_oilblock_stdev").ErrorMsg)
            SpecFailed = True
        End If

        PWM_HEAT_analogbd_stdev = BG.HeaterDT.Compute("STDEV(AnalogBdTemp)", WhereClause)
        Form1.AppendText("StDev AnalogBdTemp = " + PWM_HEAT_analogbd_stdev.ToString)
        If Not Specs("PWM_HEAT_analogbd_stdev").CompareToSpec(PWM_HEAT_analogbd_stdev) Then
            Form1.AppendText(Specs("PWM_HEAT_analogbd_stdev").ErrorMsg)
            SpecFailed = True
        End If

        Form1.AppendText("RetVal = " + BG.RetVal.ToString + vbCr)
        If BG.RetVal = False Then
            Form1.AppendText(BG.ErrorMsg)
            Return False
        End If

        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\PWM_HEAT." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
        End If

        If SpecFailed Then Return False

        Return True
    End Function

    'Shared Sub UpdateHeaterTable(ByVal status As String)
    Shared Sub UpdateHeaterTable(ByVal row As DataRow)
        BG.HeaterDT.Rows.Add(row)
        Form1.DataGridView1.FirstDisplayedCell = Form1.DataGridView1.Rows(Form1.DataGridView1.Rows.Count - 1).Cells(0)
        'Form1.DataGridView1.Refresh()
    End Sub

End Class