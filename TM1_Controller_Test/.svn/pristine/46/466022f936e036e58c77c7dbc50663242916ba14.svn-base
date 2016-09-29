Imports System.IO.Ports

Partial Class Tests
    Public Shared Function TestReworkH2ScanReady() As Boolean
        Dim Parent As Control = Form1.GetControl
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim ts_60s As New TimeSpan(0, 0, 60)
        BG = New Background(Parent, AddressOf UpdateHeaterTable)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        BG.SerialPortName = Form1.Tmcom1ComboBox.Text
        BG.NewHeaterTable()
        Form1.DataGridView1.DataSource = BG.HeaterDT
        For Each col In Form1.DataGridView1.Columns
            If Not (col.Name = "SnsrTemp" Or col.Name = "PcbTemp" Or col.Name = "OilTemp" Or
                    col.Name = "Status" Or col.Name = "Timestamp" Or col.Name = "AnalogBdTemp" Or
                    col.Name = "AnalogBdPWM") Then
                Form1.DataGridView1.Columns(col.Name).Visible = False
            Else
                Form1.DataGridView1.Columns(col.Name).Visible = True
                Form1.DataGridView1.Columns(col.Name).Width = 50
            End If
        Next

        CommonLib.Delay(5)
        If Not BG.StartLogging() Then
            Return False
        End If

        Form1.AppendText("Check for H2Scan status <> 4096.")

        Dim PassFlag As Boolean = True
        If Not BG.WaitReworkH2scanReady(600) Then
            PassFlag = False
        End If
        BG.EndLogging()

        Return PassFlag

    End Function

    Public Shared Function Test_H2SCAN_READY() As Boolean
        Dim Parent As Control = Form1.GetControl
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim startTime As DateTime
        Dim RowDateTime As DateTime
        Dim WhereClause As String
        Dim ts_60s As New TimeSpan(0, 0, 60)
        Dim PcbTempAve As Double
        Dim AnalogBdTemp As Double
        Dim passFlag As Boolean = True
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")

        BG = New Background(Parent, AddressOf UpdateHeaterTable)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        BG.SerialPortName = Form1.Tmcom1ComboBox.Text
        BG.NewHeaterTable()
        Form1.DataGridView1.DataSource = BG.HeaterDT
        For Each col In Form1.DataGridView1.Columns
            If Not (col.Name = "SnsrTemp" Or col.Name = "PcbTemp" Or col.Name = "OilTemp" Or
                    col.Name = "Status" Or col.Name = "Timestamp" Or col.Name = "AnalogBdTemp" Or
                    col.Name = "AnalogBdPWM") Then
                Form1.DataGridView1.Columns(col.Name).Visible = False
            Else
                Form1.DataGridView1.Columns(col.Name).Visible = True
                Form1.DataGridView1.Columns(col.Name).Width = 50
            End If
        Next

        CommonLib.Delay(5)
        If Not BG.StartLogging() Then
            Return False
        End If

        Form1.AppendText("Waiting for H2SCAN to be ready (status = 8001.")

        If Not BG.WaitH2scanReady(600) Then
            BG.EndLogging()
            Return False
        End If

        Form1.AppendText("CAPTURING 60s OF TM1 & HEATER DATA")
        StartTime = Now
        While BG.HeaterThread.IsAlive And Now.Subtract(StartTime).TotalSeconds < 60
            Application.DoEvents()
        End While
        BG.EndLogging()

        RowDateTime = BG.HeaterDT.Rows(BG.HeaterDT.Rows.Count - 1)("Timestamp")
        WhereClause = "Timestamp > '" + RowDateTime.Subtract(ts_60s).ToString + "'"
        PcbTempAve = BG.HeaterDT.Compute("AVG(PcbTemp)", WhereClause)
        Form1.AppendText("ave PcbTemp = " + PcbTempAve.ToString)
        AnalogBdTemp = BG.HeaterDT.Compute("AVG(AnalogBdTemp)", WhereClause)
        Form1.AppendText("ave AnalogBdTemp = " + AnalogBdTemp.ToString)

        'If Math.Abs(AnalogBdTemp - PcbTempAve) > 0.7 And Math.Abs(AnalogBdTemp + 6 - PcbTempAve) > 0.7 Then
        ' Take out the + 6 degree to intentially catch that special case
        ' Changed the spec from 0.7 to 1.0 per Thomas Walter's agreement
        ' Added back the + 6C for old H2Scan boards.
        If Math.Abs(AnalogBdTemp - PcbTempAve) > 1.0 And Math.Abs(AnalogBdTemp + 6 - PcbTempAve) > 1.0 Then
            Form1.AppendText("Expected diff between AnalogBdTemp and PcbTemp < 1.0")
            passFlag = False
        End If

        ' Save the data in .csv format
        csv_filepath = ReportDir + Form1.AssemblyLevel_ComboBox.Text + "\" + Form1.Serial_Number + "\H2SCAN_READY." + TimeStamp + ".csv"
        If Not CommonLib.ExportDataTableToCSV(BG.HeaterDT, csv_filepath) Then
            Form1.AppendText("Problem creating csv file " + csv_filepath)
            Return False
        End If

        Return passFlag
    End Function
End Class