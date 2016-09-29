Imports System.IO.Ports

Partial Class Tests
    Public Shared Function Test_TEC_COOL() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Parent As Control = Form1.GetControl
        Dim NewConfig As New Hashtable
        Dim RebootNeeded As Boolean
        Dim T As New TM1
        Dim AnalogBdTemp_start As Double
        Dim OilBlockTemp_start As Double
        Dim startTime As DateTime
        Dim last_row As Integer
        Dim AnalogBd_done As Boolean
        Dim OilBlock_done As Boolean
        Dim AnalogBd_cool_time As Integer
        Dim OilBlock_cool_time As Integer
        Dim Passed As Boolean = False
        Dim Temp_fall_timeout As Integer = 370
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim SpecFailed As Boolean
        Dim Response As String
        Dim conditions As Hashtable
        Dim cond As condition

        BG = New Background(Parent, AddressOf UpdateHeaterTable)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If


        NewConfig.Add("CLI_OVER_TMCOM1.ENABLE", "TRUE")
        NewConfig.Add("TMCOM1.PROTOCOL", "CLI")
        NewConfig.Add("TMCOM1.MODE", "RS232")
        If Not T.SetConfig(NewConfig, RebootNeeded, True) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "sensor off", 5) Then
            Form1.AppendText("Problem sending command 'sensor off'")
            Return False
        End If
        CommonLib.Delay(10)

        BG.SerialPortName = Form1.Tmcom1ComboBox.Text
        BG.NewHeaterTable()
        Form1.DataGridView1.DataSource = BG.HeaterDT
        For Each col In Form1.DataGridView1.Columns
            If Not (col.Name = "Timestamp" Or col.Name = "OilBlockTemp" Or col.Name = "AnalogBdTemp") Then
                Form1.DataGridView1.Columns(col.Name).Visible = False
            Else
                Form1.DataGridView1.Columns(col.Name).Visible = True
                Form1.DataGridView1.Columns(col.Name).Width = 50
            End If
        Next

        BG.StartLogging()
        startTime = Now
        While (BG.HeaterDT.Rows.Count = 0 And Now.Subtract(startTime).TotalSeconds < 10)
            Application.DoEvents()
        End While
        If BG.HeaterDT.Rows.Count = 0 Then
            Form1.AppendText("Heater data capture did not start")
            T.SetTEC_AUTO()
            BG.EndLogging()
            Return False
        End If

        conditions = New Hashtable

        cond = New condition
        cond.expected = 50.0
        conditions.Add("AnalogBdTemp", cond)

        cond = New condition
        cond.expected = 45.0
        conditions.Add("OilBlockTemp", cond)

        If Not BG.GetZonesToTemp(conditions, 1000, "TEC_COOL") Then
            Form1.AppendText("Zones did not reach initial target temperatures")
            BG.EndLogging()
            Return False
        End If

        If Not T.SetPWM("OILBLOCK", 0) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        If Not T.SetPWM("ANALOGBD", 0) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        If Not T.SetTEC("OILBLOCK", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        If Not T.SetTEC("ANALOGBD", "COOL", 0) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        CommonLib.Delay(20)
        If Not T.SetTEC("OILBLOCK", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        If Not T.SetTEC("ANALOGBD", "COOL", 100) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        AnalogBdTemp_start = BG.HeaterDT.Rows(BG.HeaterDT.Rows.Count - 1)("AnalogBdTemp")
        OilBlockTemp_start = BG.HeaterDT.Rows(BG.HeaterDT.Rows.Count - 1)("OilBlockTemp")
        Form1.AppendText("Waiting for 'AnalogBdTemp' to reach " + (AnalogBdTemp_start - 5).ToString)
        Form1.AppendText("Waiting for 'OilBlockTemp' to reach " + (OilBlockTemp_start - 5).ToString)

        AnalogBd_done = False
        OilBlock_done = False
        startTime = Now
        While ((Not AnalogBd_done Or Not OilBlock_done) And Now.Subtract(startTime).TotalSeconds < Temp_fall_timeout)
            Application.DoEvents()
            last_row = BG.HeaterDT.Rows.Count - 1
            If Not AnalogBd_done And BG.HeaterDT.Rows(last_row)("AnalogBdTemp") < (AnalogBdTemp_start - 5) Then
                AnalogBd_cool_time = Now.Subtract(startTime).TotalSeconds
                Form1.AppendText("AnalogBd 5C temp drop completed in " + AnalogBd_cool_time.ToString + " seconds")
                AnalogBd_done = True
                If Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                    Form1.AppendText(T.ErrorMsg)
                    BG.EndLogging()
                    Return False
                End If
            End If
            If Not OilBlock_done And BG.HeaterDT.Rows(last_row)("OilBlockTemp") < (OilBlockTemp_start - 5) Then
                OilBlock_cool_time = Now.Subtract(startTime).TotalSeconds
                Form1.AppendText("OilBlock 5C temp drop completed in " + OilBlock_cool_time.ToString + " seconds")
                OilBlock_done = True
                If Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                    Form1.AppendText(T.ErrorMsg)
                    BG.EndLogging()
                    Return False
                End If
            End If
            Form1.TimeoutLabel.Text = Math.Round(Temp_fall_timeout - Now.Subtract(startTime).TotalSeconds).ToString
        End While
        BG.EndLogging()
        If Not T.SetTEC_AUTO() Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        If Not AnalogBd_done Then
            Form1.AppendText("Timeout waiting for AnalogBd temperature to fall 5C")
        End If
        If Not OilBlock_done Then
            Form1.AppendText("Timeout waiting for OilBlock temperature to fall 5C")
        End If
        Passed = AnalogBd_done And OilBlock_done

        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\TEC_COOL." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
        End If

        SpecFailed = False
        If Not Specs("TEC_COOL_oilblock_time").CompareToSpec(OilBlock_cool_time) Then
            Form1.AppendText(Specs("TEC_COOL_oilblock_time").ErrorMsg)
            SpecFailed = True
        End If
        If Not Specs("TEC_COOL_analogbd_time").CompareToSpec(AnalogBd_cool_time) Then
            Form1.AppendText(Specs("TEC_COOL_analogbd_time").ErrorMsg)
            SpecFailed = True
        End If

        If Not SF.Cmd(SerialPort, Response, "sensor on", 5) Then
            Form1.AppendText("Problem sending command 'sensor off'")
            Return False
        End If

        If SpecFailed Then Return False

        Return Passed
    End Function
End Class