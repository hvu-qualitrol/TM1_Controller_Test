Imports System.IO.Ports

Partial Class Tests
    Shared LF As Leak_Fixture

    Public Shared Function Test_LEAK() As Boolean
        Dim row_cnt As Integer
        Dim startTime As DateTime
        Dim Parent As Control = Form1.GetControl
        'Dim LF As New Leak_Fixture(Parent, AddressOf UpdatePressureTable)
        Dim MIN_STARTING_PRESSURE = 80
        ' Dim LEAK_TEST_TIME As Integer = 20 * 60
        Dim LEAK_TEST_TIME As Integer = 7 * 60
        Dim start, finish As Double
        Dim slope As Double
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")

        LF = New Leak_Fixture(Parent, AddressOf UpdatePressureTable)

        LF.SerialPortName = Form1.Tmcom1ComboBox.Text
        LF.NewPressureTable()
        Form1.DataGridView1.DataSource = LF.PressureDT
        'Form1.DataGridView1.Columns("Pressure").DefaultCellStyle.Format() = "n2"

        row_cnt = 0
        LF.StartLogging()

        Form1.AppendText("PURGING FIXTURE")
        If Not LF.ChangeValves(15) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        'CommonLib.Delay(10)
        CommonLib.Delay(20)

        Form1.AppendText("CLOSING OUTLET VALVES")
        If Not LF.ChangeValves(3) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If

        'CommonLib.Delay(5)
        CommonLib.Delay(10)
        If (LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("Pressure") < MIN_STARTING_PRESSURE) Then
            Form1.AppendText("Expected pressure after closing the outlet vale >= " + MIN_STARTING_PRESSURE.ToString)
            CleanupOnAbort()
            Return False
        End If

        Form1.AppendText("CLOSING INLET VALVES")
        If Not LF.ChangeValves(0) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        'CommonLib.Delay(15)
        CommonLib.Delay(20)

        'Wait 5 minutes for the pressure to build up and stablelize
        Form1.AppendText("Wait 5 minutes for the pressure to build up and stablelize...")
        CommonLib.Delay(60 * 5)

        Form1.AppendText("CAPTURING PRESSURE DATA")
        start = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        startTime = Now
        Dim CurrentRowCount As Integer = LF.PressureDT.Rows.Count - 1
        Dim LastRowCount As Integer = CurrentRowCount
        'Dim Min As Double = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("Pressure")
        'Dim Max As Double = Min
        While LF.LeakThread.IsAlive And Now.Subtract(startTime).TotalSeconds < LEAK_TEST_TIME
            Application.DoEvents()
            If Stopped Then
                CleanupOnAbort()
                Return False
            Else
                CurrentRowCount = LF.PressureDT.Rows.Count - 1
                If CurrentRowCount > LastRowCount Then
                    LastRowCount = CurrentRowCount
                    If LF.PressureDT(CurrentRowCount)("Pressure") < MIN_STARTING_PRESSURE Then
                        Form1.AppendText("Expected pressure >= " + MIN_STARTING_PRESSURE.ToString)
                        CleanupOnAbort()
                    End If
                    '    If LF.PressureDT(CurrentRowCount)("Pressure") > Max Then
                    '        Max = LF.PressureDT(CurrentRowCount)("Pressure")
                    '    End If
                    '    If LF.PressureDT(CurrentRowCount)("Pressure") < Min Then
                    '        Min = LF.PressureDT(CurrentRowCount)("Pressure")
                    '    End If
                    '    If Math.Abs((Max - Min) / ((Max + Min) / 2)) > 0.05 Then
                    '        Form1.AppendText("Expected variation % valve >= 5%")
                    '        CleanupOnAbort()
                    '        Return False
                    '    End If
                End If
            End If
            Form1.TimeoutLabel.Text = (Math.Round(LEAK_TEST_TIME - Now.Subtract(startTime).TotalSeconds)).ToString
        End While
        Form1.AppendText("ENDING PRESSURE DATA CAPTURE")
        finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        If Not LF.ChangeValves(12) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        LF.StopLogging = True
        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\LEAK." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(LF.PressureDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
        End If
        CommonLib.Delay(5)
        'finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        'Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        'If Not LF.ChangeValves(12) Then
        '    Form1.AppendText(LF.ErrorMsg)
        '    If LF.LeakThread.IsAlive Then
        '        LF.StopLogging = True
        '        CommonLib.Delay(5)
        '        LF.LeakThread.Abort()
        '        LF.LeakThread.Join()
        '    End If
        '    Return False
        'End If
        LF.LeakThread.Abort()
        LF.LeakThread.Join()
        'finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")

        'CommonLib.Delay(10)
        'Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        'If Not LF.ChangeValves(12) Then
        '    Form1.AppendText(LF.ErrorMsg)
        '    If LF.LeakThread.IsAlive Then
        '        LF.StopLogging = True
        '        CommonLib.Delay(5)
        '        LF.LeakThread.Abort()
        '        LF.LeakThread.Join()
        '    End If
        '    Return False
        'End If

        Form1.AppendText("Calculating leakrate from " + start.ToString + " to " + finish.ToString)
        'Dim LeakStdevSpec As Double = Convert.ToDouble(Form1.TextBoxLeakStdev.Text.ToString)
        'Dim LeakRateSpec As Double = Convert.ToDouble(Form1.TextBoxLeakRateSpec.Text.ToString)
        'Form1.AppendText("Leak rate spec = " + LeakRateSpec.ToString + " Leak stdev spec = " + LeakStdevSpec.ToString)

        Dim PassFlag As Boolean = True
        Dim LeakStdevSpec As Double = 0.1
        Dim stdev As Double = LF.CalculatePressureStdev(start, finish)
        Form1.AppendText("stdev = " + stdev.ToString + " psi/min.")
        If stdev > LeakStdevSpec Then
            Form1.AppendText("Failed: Standard deviation exceed limit " + LeakStdevSpec.ToString)
            PassFlag = False
        End If

        Dim LeakRateSpec As Double = -0.02
        slope = LF.CalculateLeakRate2(start, finish)
        Form1.AppendText("leak rate = " + slope.ToString + " psi/min")
        If slope < LeakRateSpec Then
            Form1.AppendText("Failed: Leak rate exceed limit " + LeakRateSpec.ToString)
            PassFlag = False
        End If

        Return PassFlag

        'If Not Specs("Leak_Rate").CompareToSpec(slope) Then
        '    Form1.AppendText(Specs("Leak_Rate").ErrorMsg)
        '    Return False
        'End If

        'Return True

    End Function

    Public Shared Function Test_LEAK0() As Boolean
        Dim row_cnt As Integer
        Dim startTime As DateTime
        Dim Parent As Control = Form1.GetControl
        'Dim LF As New Leak_Fixture(Parent, AddressOf UpdatePressureTable)
        Dim MIN_STARTING_PRESSURE = 80
        ' Dim LEAK_TEST_TIME As Integer = 20 * 60
        Dim LEAK_TEST_TIME As Integer = 7 * 60
        Dim start, finish As Double
        Dim slope As Double
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")

        LF = New Leak_Fixture(Parent, AddressOf UpdatePressureTable)

        LF.SerialPortName = Form1.Tmcom1ComboBox.Text
        LF.NewPressureTable()
        Form1.DataGridView1.DataSource = LF.PressureDT
        'Form1.DataGridView1.Columns("Pressure").DefaultCellStyle.Format() = "n2"

        row_cnt = 0
        LF.StartLogging()

        Form1.AppendText("PURGING FIXTURE")
        If Not LF.ChangeValves(15) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        'CommonLib.Delay(10)
        CommonLib.Delay(20)

        Form1.AppendText("CLOSING OUTLET VALVES")
        If Not LF.ChangeValves(3) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If

        'CommonLib.Delay(5)
        CommonLib.Delay(10)
        If (LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("Pressure") < MIN_STARTING_PRESSURE) Then
            Form1.AppendText("Expected pressure after closing the outlet vale >= " + MIN_STARTING_PRESSURE.ToString)
            CleanupOnAbort()
            Return False
        End If

        Form1.AppendText("CLOSING INLET VALVES")
        If Not LF.ChangeValves(0) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        'CommonLib.Delay(15)
        CommonLib.Delay(20)

        Form1.AppendText("CAPTURING PRESSURE DATA")
        start = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        startTime = Now
        Dim CurrentRowCount As Integer = LF.PressureDT.Rows.Count - 1
        Dim LastRowCount As Integer = CurrentRowCount
        Dim Min As Double = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("Pressure")
        Dim Max As Double = Min
        While LF.LeakThread.IsAlive And Now.Subtract(startTime).TotalSeconds < LEAK_TEST_TIME
            Application.DoEvents()
            If Stopped Then
                CleanupOnAbort()
                Return False
            Else
                CurrentRowCount = LF.PressureDT.Rows.Count - 1
                If CurrentRowCount > LastRowCount Then
                    LastRowCount = CurrentRowCount
                    If LF.PressureDT(CurrentRowCount)("Pressure") < MIN_STARTING_PRESSURE Then
                        Form1.AppendText("Expected pressure >= " + MIN_STARTING_PRESSURE.ToString)
                        CleanupOnAbort()
                    End If
                    If LF.PressureDT(CurrentRowCount)("Pressure") > Max Then
                        Max = LF.PressureDT(CurrentRowCount)("Pressure")
                    End If
                    If LF.PressureDT(CurrentRowCount)("Pressure") < Min Then
                        Min = LF.PressureDT(CurrentRowCount)("Pressure")
                    End If
                    If Math.Abs((Max - Min) / ((Max + Min) / 2)) > 0.05 Then
                        Form1.AppendText("Expected variation % valve >= 5%")
                        CleanupOnAbort()
                        Return False
                    End If
                End If
            End If
            Form1.TimeoutLabel.Text = (Math.Round(LEAK_TEST_TIME - Now.Subtract(startTime).TotalSeconds)).ToString
        End While
        Form1.AppendText("ENDING PRESSURE DATA CAPTURE")
        finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        If Not LF.ChangeValves(12) Then
            Form1.AppendText(LF.ErrorMsg)
            CleanupOnAbort()
            Return False
        End If
        LF.StopLogging = True
        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\LEAK." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(LF.PressureDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
        End If
        CommonLib.Delay(5)
        'finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")
        'Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        'If Not LF.ChangeValves(12) Then
        '    Form1.AppendText(LF.ErrorMsg)
        '    If LF.LeakThread.IsAlive Then
        '        LF.StopLogging = True
        '        CommonLib.Delay(5)
        '        LF.LeakThread.Abort()
        '        LF.LeakThread.Join()
        '    End If
        '    Return False
        'End If
        LF.LeakThread.Abort()
        LF.LeakThread.Join()
        'finish = LF.PressureDT.Rows(LF.PressureDT.Rows.Count - 1)("TimeSinceStart")

        'CommonLib.Delay(10)
        'Form1.AppendText("RELEASING PRESSURE FROM FIXTURE")
        'If Not LF.ChangeValves(12) Then
        '    Form1.AppendText(LF.ErrorMsg)
        '    If LF.LeakThread.IsAlive Then
        '        LF.StopLogging = True
        '        CommonLib.Delay(5)
        '        LF.LeakThread.Abort()
        '        LF.LeakThread.Join()
        '    End If
        '    Return False
        'End If

        Form1.AppendText("Calculating leakrate from " + start.ToString + " to " + finish.ToString)
        'Dim LeakStdevSpec As Double = Convert.ToDouble(Form1.TextBoxLeakStdev.Text.ToString)
        'Dim LeakRateSpec As Double = Convert.ToDouble(Form1.TextBoxLeakRateSpec.Text.ToString)
        'Form1.AppendText("Leak rate spec = " + LeakRateSpec.ToString + " Leak stdev spec = " + LeakStdevSpec.ToString)

        Dim PassFlag As Boolean = True
        Dim LeakStdevSpec As Double = 0.1
        Dim stdev As Double = LF.CalculatePressureStdev(start, finish)
        Form1.AppendText("stdev = " + stdev.ToString + " psi/min.")
        If stdev > LeakStdevSpec Then
            Form1.AppendText("Failed: Standard deviation exceed limit " + LeakStdevSpec.ToString)
            PassFlag = False
        End If

        Dim LeakRateSpec As Double = -0.0145
        slope = LF.CalculateLeakRate2(start, finish)
        Form1.AppendText("leak rate = " + slope.ToString + " psi/min")
        If slope < LeakRateSpec Then
            Form1.AppendText("Failed: Leak rate exceed limit " + LeakRateSpec.ToString)
            PassFlag = False
        End If

        Return PassFlag

        'If Not Specs("Leak_Rate").CompareToSpec(slope) Then
        '    Form1.AppendText(Specs("Leak_Rate").ErrorMsg)
        '    Return False
        'End If

        'Return True

    End Function

    'Shared Sub UpdatePressureTable(ByVal status As String)
    Shared Sub UpdatePressureTable(ByVal row As DataRow)
        LF.PressureDT.Rows.Add(row)
        'Form1.DataGridView1.Columns.Item(3).DefaultCellStyle.Format()
        'If Form1.DataGridView1.Rows.Count = 1 Then
        '    Form1.DataGridView1.Columns("Pressure").DefaultCellStyle.Format() = "n2"
        'End If
        Form1.DataGridView1.FirstDisplayedCell = Form1.DataGridView1.Rows(Form1.DataGridView1.Rows.Count - 1).Cells(0)
        'Application.DoEvents()
        'If Stopped Then
        '    CleanupOnAbort()
        'End If
        'Form1.DataGridView1.Refresh() 
    End Sub

    ' Clean up on abort
    Shared Sub CleanupOnAbort()
        Form1.AppendText("CleanupOnAbort(): " + LF.ErrorMsg)
        If LF.LeakThread.IsAlive Then
            'LF.ChangeValves(12)
            LF.StopLogging = True
            CommonLib.Delay(5)
            'LF.LeakThread.Abort()
            LF.LeakThread.Join()
            Form1.AppendText("LeakThread aborted. Last error msg: " + LF.ErrorMsg)
        End If
    End Sub

End Class