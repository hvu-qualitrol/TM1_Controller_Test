Imports System.IO.Ports
Imports System.Text.RegularExpressions

Public Class Background
    Private MAX_TEMP As Double = 75
    Friend SerialPortName As String
    Friend RetVal As Boolean
    'Friend HeaterData As Queue
    Friend HeaterDT As DataTable
    Friend ErrorMsg As String = "UNKNOWN"
    Friend StopLogging As Boolean = False
    Friend ClearDT As Boolean = False
    Public HeaterThread As System.Threading.Thread
    Dim results As ReturnResults

    Private m_BaseControl As Control
    Private m_CallBackFunction As CallBackDelegate

    Sub New(ByRef caller As Control, ByRef callbackFunction As CallBackDelegate)
        m_BaseControl = caller
        m_CallBackFunction = callbackFunction
    End Sub

    Function StartLogging() As Boolean
        Dim startTime As DateTime = Now

        HeaterThread = New System.Threading.Thread(AddressOf LogHeater)
        HeaterThread.Start()
        While (HeaterDT.Rows.Count = 0 And Now.Subtract(startTime).TotalSeconds < 15)
            Application.DoEvents()
        End While

        If HeaterDT.Rows.Count = 0 Then
            Form1.AppendText("Data logging failed to start")
            If HeaterThread.IsAlive Then
                EndLogging()
            Else
                If Not RetVal Then
                    Form1.AppendText(ErrorMsg)
                End If
            End If
            Return False
        End If

        Form1.AppendText("Data logging started in " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString)

        Return True
    End Function

    Sub EndLogging()
        StopLogging = True
        CommonLib.Delay(10)
        ' CommonLib.Delay(5)
        'HeaterThread.Abort()
        HeaterThread.Join()
        Form1.AppendText("EndLogging(): Heater thread was killed.")
    End Sub

    Sub LogHeater()
        Dim SF As New SerialFunctions
        Dim LineCnt As Integer
        Dim Line As String
        Dim Fields() As String
        Dim DataLineFields() As String = {"", "H2_OIL|I", "H2_DGA.PPM|I", "H2.PPM|I", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "Status|I",
"HCurrent|I", "ResAdc|I", "AdjRes|I", "H2Res.PPM|I", "OilBlockTemp|D", "OilBlockPWM|D", "OilBlockDir|S", "AnalogBdTemp|D", "AnalogBdPWM|D",
"AnalogBdDir|S", "AmbientTemp|D", "OilblockRevPWM|D", "AnalogBdRevPWM|D", "Tach|I"}
        Dim Name As String
        Dim row As DataRow
        Dim TypeStr As String

        Tmcom1SerialPort = New SerialPort(SerialPortName, 115200, 0, 8, 1)
        Tmcom1SerialPort.Handshake = Handshake.RequestToSend
        Try
            Tmcom1SerialPort.Open()
        Catch ex As Exception
            ErrorMsg = "Problem opening TMCOM1 serial port" + vbCr
            ErrorMsg += ex.ToString
            RetVal = False
            Exit Sub
        End Try
        Tmcom1SerialPort.ReadTimeout = 2000
        Tmcom1SerialPort.NewLine = Chr(10)
        Tmcom1SerialPort.Write(Chr(4))
        System.Threading.Thread.Sleep(200)
        results = SF.Login(Tmcom1SerialPort)
        If Not results.PassFail Then
            ErrorMsg = "Login failed" + vbCr
            ErrorMsg += results.Result + vbCr
            Tmcom1SerialPort.Close()
            RetVal = False
            Exit Sub
        End If

        LineCnt = 0
        System.Threading.Thread.Sleep(200)
        'CommonLib.Delay(5)
        Tmcom1SerialPort.Write("heater" + Chr(10))
        LineCnt = 0
        'While (LineCnt < 60)
        While (Not StopLogging)
            If ClearDT Then
                HeaterDT.Rows.Clear()
                ClearDT = False
            End If
            Try
                Line = Tmcom1SerialPort.ReadLine()
            Catch ex As Exception
                Line = ""
            End Try
            If Regex.IsMatch(Line, "^\d\d\d\d-") Then
                Line = Regex.Replace(Line, Chr(13), "")
                Line = Regex.Replace(Line, Chr(10), "")
                Fields = Split(Line, ",")
                If Not Fields.Count = DataLineFields.Count Then
                    ErrorMsg = "Heater data line has " + Fields.Count.ToString + " fields, expecting " + DataLineFields.Count.ToString
                    RetVal = False
                    Exit Sub
                End If
                'HeaterData.Enqueue(Line)
                row = HeaterDT.NewRow()

                For i = 0 To UBound(HeaterFields)
                    Name = Split(HeaterFields(i), "|")(0)
                    TypeStr = Split(HeaterFields(i), "|")(1)
                    Select Case TypeStr
                        Case "DT"
                            row(Name) = DateTime.Parse(Fields(i))
                        Case "D"
                            row(Name) = CDbl(Fields(i))
                        Case "I"
                            row(Name) = CInt(Fields(i))
                        Case "S"
                            row(Name) = Fields(i).ToString
                        Case Else
                            ErrorMsg = "No conversion for heater field " + Name + " with type " + TypeStr
                            RetVal = False
                            Exit Sub
                    End Select
                    'row(Name) = Fields(i)
                Next
                m_BaseControl.Invoke(m_CallBackFunction, row)
                'HeaterDT.Rows.Add(row)
                LineCnt += 1
            End If
        End While

        Tmcom1SerialPort.Close()
        Form1.AppendText("LogHeater() ended: Tmcom1SerialPort.Close().")

        RetVal = True
    End Sub

    Sub NewHeaterTable()
        Dim Name As String
        Dim TypeStr As String
        Dim Field As String

        HeaterDT = New DataTable()

        For Each Field In HeaterFields
            Name = Split(Field, "|")(0)
            TypeStr = Split(Field, "|")(1)
            Select Case TypeStr
                Case "DT"
                    HeaterDT.Columns.Add(Name, Type.GetType("System.DateTime"))
                Case "D"
                    HeaterDT.Columns.Add(Name, Type.GetType("System.Double"))
                Case "I"
                    HeaterDT.Columns.Add(Name, Type.GetType("System.Int32"))
                Case "S"
                    HeaterDT.Columns.Add(Name, Type.GetType("System.String"))
            End Select
        Next
        ' AddHandler HeaterDT.RowChanged, New DataRowChangeEventHandler(AddressOf NewHeaterData)
    End Sub

    Function CheckInitConditions(ByRef conditions As Hashtable)
        Dim row_cnt As Integer
        Dim T As New TM1

        Form1.AppendText("-----------------------------------------------------")

        ' Check the initial temp of AnalogBdTemp against the limit
        row_cnt = HeaterDT.Rows.Count
        Dim testTemp As Double = HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp")
        If testTemp > conditions("AnalogBdTemp").expected Then
            Form1.AppendText("Failed: Initial AnalogBdTemp " + testTemp.ToString + " > " + conditions("AnalogBdTemp").expected.ToString)
            Return False
        Else
            Form1.AppendText("Initial AnalogBdTemp = " + testTemp.ToString)
        End If

        ' Check the initial temp of OilBlockTemp against the limit
        testTemp = HeaterDT.Rows(row_cnt - 1)("OilBlockTemp")
        If testTemp > conditions("OilBlockTemp").expected Then
            Form1.AppendText("Failed: Initial OilBlockTemp " + testTemp.ToString + " > " + conditions("OilBlockTemp").expected.ToString)
            Return False
        Else
            Form1.AppendText("Initial OilBlockTemp = " + testTemp.ToString)
        End If

        Return True

    End Function

    Function GetReworkZonesToTemp(ByVal HighTest As Boolean, ByVal Timeout As Integer)
        Dim conditions_met As Boolean = False
        Dim max_temp_exceeded As Boolean = False
        Dim startTime As DateTime
        Dim row_cnt As Integer
        Dim NewHeaterDateTime As DateTime
        Dim TempSensors() As String = {"SnsrTemp", "PcbTemp", "OilTemp", "OilBlockTemp", "AnalogBdTemp"}
        Dim TempSensor As String
        Dim condition_name As String
        Dim expected As Double
        Dim OilBlockControlsSet As Boolean = False
        Dim AnalogBdControlsSet As Boolean = False
        Dim T As New TM1

        Form1.AppendText("-----------------------------------------------------")

        ' Set the delta temp accordingly to the test type
        Dim DeltaTemp As Double = 3.0
        If Not HighTest Then
            DeltaTemp = -3.0
        End If

        ' Establish the test conditions based on the current temps and delta temp
        row_cnt = HeaterDT.Rows.Count
        Dim conditions As Hashtable = New Hashtable
        Dim cond As condition = New condition
        conditions.Add("AnalogBdTemp", cond)
        cond.expected = HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp") + DeltaTemp
        cond = New condition
        conditions.Add("OilBlockTemp", cond)
        cond.expected = HeaterDT.Rows(row_cnt - 1)("OilBlockTemp") + DeltaTemp

        For Each condition_name In conditions.Keys
            If condition_name = "OilBlockTemp" Then
                If Math.Abs(HeaterDT.Rows(row_cnt - 1)("OilBlockTemp") - conditions(condition_name).expected) <= 0.1 Then
                    If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).met = True
                    Form1.AppendText("OilBlockTemp temperature already met")
                ElseIf HeaterDT.Rows(row_cnt - 1)("OilBlockTemp") > conditions(condition_name).expected Then
                    Form1.AppendText("Cooling OILBLOCK")
                    If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "COOL", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).check = "<="
                Else
                    Form1.AppendText("Heating OILBLOCK")
                    If Not T.SetPWM("OILBLOCK", 100) Or Not T.SetTEC("OILBLOCK", "HEAT", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).check = ">="
                End If
                OilBlockControlsSet = True
            End If
            If condition_name = "AnalogBdTemp" Then
                If Math.Abs(HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp") - conditions(condition_name).expected) <= 0.1 Then
                    If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).met = True
                    Form1.AppendText("AnalogBd temperature already met met")
                ElseIf HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp") > conditions(condition_name).expected Then
                    If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "COOL", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    Form1.AppendText("Cooling AnalogBd")
                    conditions(condition_name).check = "<="
                Else
                    If Not T.SetPWM("ANALOGBD", 100) Or Not T.SetTEC("ANALOGBD", "HEAT", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    Form1.AppendText("Heating AnalogBd")
                    conditions(condition_name).check = ">="
                End If
                AnalogBdControlsSet = True
            End If
        Next
        If Not OilBlockControlsSet Then
            If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
        End If
        If Not AnalogBdControlsSet Then
            If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
        End If


        ' get last row
        ' verify last row not old
        startTime = Now
        While (Not max_temp_exceeded And Not conditions_met And Now.Subtract(startTime).TotalSeconds < Timeout)
            Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                'row_cnt = HeaterDT.Rows.Count
                NewHeaterDateTime = HeaterDT.Rows(row_cnt - 1)("Timestamp")
                For Each TempSensor In TempSensors
                    If HeaterDT.Rows(row_cnt - 1)(TempSensor) > MAX_TEMP Then
                        Form1.AppendText(TempSensor + " Exceeded " + MAX_TEMP.ToString)
                        max_temp_exceeded = True
                        T.SetPWM("OILBLOCK", 0)
                        T.SetTEC("OILBLOCK", "HEAT", 0)
                        T.SetPWM("ANALOGBD", 0)
                        T.SetTEC("ANALOGBD", "HEAT", 0)
                        Exit While
                    End If
                Next
                conditions_met = True
                For Each condition_name In conditions.Keys
                    If conditions(condition_name).met Then Continue For
                    conditions_met = False
                    expected = conditions(condition_name).expected
                    If (conditions(condition_name).check = ">=" And HeaterDT.Rows(row_cnt - 1)(condition_name) > expected) Or
                        (conditions(condition_name).check = "<=" And HeaterDT.Rows(row_cnt - 1)(condition_name) < expected) Then
                        Form1.AppendText(condition_name + " " + HeaterDT.Rows(row_cnt - 1)(condition_name).ToString + " has reached " + expected.ToString, False)
                        conditions(condition_name).met = True
                        If condition_name = "OilBlockTemp" Then
                            If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                                Form1.AppendText(T.ErrorMsg)
                                Return False
                            End If
                        End If
                        If condition_name = "AnalogBdTemp" Then
                            If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                                Form1.AppendText(T.ErrorMsg)
                                Return False
                            End If
                        End If
                        'Else
                        '    Form1.AppendText(condition_name + " " + HeaterDT.Rows(row_cnt - 1)(condition_name).ToString + " has not reached " + expected.ToString, False)
                    End If
                Next
                row_cnt += 1
            ElseIf Now.Subtract(HeaterDT.Rows(row_cnt - 1)("Timestamp")).TotalSeconds > 10 Then
                Form1.AppendText("No heater data in 10s")
                Exit While
            End If
        End While

        T.SetPWM("OILBLOCK", 0)
        T.SetTEC("OILBLOCK", "HEAT", 0)
        T.SetPWM("ANALOGBD", 0)
        T.SetTEC("ANALOGBD", "HEAT", 0)

        If conditions_met Then
            Form1.AppendText("zones reach expected values in " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString)
        Else
            For Each condition_name In conditions.Keys
                If Not conditions(condition_name).met Then
                    Form1.AppendText("condition " + condition_name + " not met")
                End If
            Next
        End If
        Form1.AppendText("-----------------------------------------------------")

        Return conditions_met
    End Function


    'Sub NewHeaterData(ByVal sender As Object, ByVal e As DataRowChangeEventArgs)
    '    m_BaseControl.Invoke(m_CallBackFunction, "HERE")
    'End Sub

    Function GetZonesToTemp(ByRef conditions As Hashtable, ByVal Timeout As Integer, Optional ByVal testName As String = "Test")
        Dim conditions_met As Boolean = False
        Dim max_temp_exceeded As Boolean = False
        Dim startTime As DateTime
        Dim row_cnt As Integer
        Dim NewHeaterDateTime As DateTime
        Dim TempSensors() As String = {"SnsrTemp", "PcbTemp", "OilTemp", "OilBlockTemp", "AnalogBdTemp"}
        Dim TempSensor As String
        Dim condition_name As String
        Dim limit As Double
        Dim expected As Double
        Dim OilBlockControlsSet As Boolean = False
        Dim AnalogBdControlsSet As Boolean = False
        Dim T As New TM1

        Form1.AppendText("-----------------------------------------------------")
        For Each condition_name In conditions.Keys
            Form1.AppendText("Driving " + condition_name + " to " + conditions(condition_name).Expected.ToString)
        Next

        row_cnt = HeaterDT.Rows.Count
        For Each condition_name In conditions.Keys
            If condition_name = "OilBlockTemp" Then
                If Math.Abs(HeaterDT.Rows(row_cnt - 1)("OilBlockTemp") - conditions(condition_name).expected) <= 0.1 Then
                    If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).met = True
                    Form1.AppendText("OilBlockTemp temperature already met")
                ElseIf HeaterDT.Rows(row_cnt - 1)("OilBlockTemp") > conditions(condition_name).expected Then
                    Form1.AppendText("Cooling OILBLOCK")
                    If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "COOL", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).check = "<="
                Else
                    Form1.AppendText("Heating OILBLOCK")
                    If Not T.SetPWM("OILBLOCK", 100) Or Not T.SetTEC("OILBLOCK", "HEAT", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).check = ">="
                End If
                OilBlockControlsSet = True
            End If
            If condition_name = "AnalogBdTemp" Then
                If Math.Abs(HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp") - conditions(condition_name).expected) <= 0.1 Then
                    If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    conditions(condition_name).met = True
                    Form1.AppendText("AnalogBd temperature already met met")
                ElseIf HeaterDT.Rows(row_cnt - 1)("AnalogBdTemp") > conditions(condition_name).expected Then
                    If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "COOL", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    Form1.AppendText("Cooling AnalogBd")
                    conditions(condition_name).check = "<="
                Else
                    If Not T.SetPWM("ANALOGBD", 100) Or Not T.SetTEC("ANALOGBD", "HEAT", 100) Then
                        Form1.AppendText(T.ErrorMsg)
                        Return False
                    End If
                    Form1.AppendText("Heating AnalogBd")
                    conditions(condition_name).check = ">="
                End If
                AnalogBdControlsSet = True
            End If
        Next
        If Not OilBlockControlsSet Then
            If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
        End If
        If Not AnalogBdControlsSet Then
            If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
        End If


        ' get last row
        ' verify last row not old
        startTime = Now
        While (Not max_temp_exceeded And Not conditions_met And Now.Subtract(startTime).TotalSeconds < Timeout)
            Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                'row_cnt = HeaterDT.Rows.Count
                NewHeaterDateTime = HeaterDT.Rows(row_cnt - 1)("Timestamp")
                For Each TempSensor In TempSensors
                    If HeaterDT.Rows(row_cnt - 1)(TempSensor) > MAX_TEMP Then
                        Form1.AppendText(TempSensor + " Exceeded " + MAX_TEMP.ToString)
                        max_temp_exceeded = True
                        T.SetPWM("OILBLOCK", 0)
                        T.SetTEC("OILBLOCK", "HEAT", 0)
                        T.SetPWM("ANALOGBD", 0)
                        T.SetTEC("ANALOGBD", "HEAT", 0)
                        Exit While
                    End If
                Next
                conditions_met = True
                For Each condition_name In conditions.Keys
                    If conditions(condition_name).met Then Continue For
                    conditions_met = False
                    expected = conditions(condition_name).expected
                    If (conditions(condition_name).check = ">=" And HeaterDT.Rows(row_cnt - 1)(condition_name) > expected) Or
                        (conditions(condition_name).check = "<=" And HeaterDT.Rows(row_cnt - 1)(condition_name) < expected) Then
                        Form1.AppendText(condition_name + " " + HeaterDT.Rows(row_cnt - 1)(condition_name).ToString + " has reached " + expected.ToString)
                        conditions(condition_name).met = True
                        If condition_name = "OilBlockTemp" Then
                            If Not T.SetPWM("OILBLOCK", 0) Or Not T.SetTEC("OILBLOCK", "HEAT", 0) Then
                                Form1.AppendText(T.ErrorMsg)
                                Return False
                            End If
                        End If
                        If condition_name = "AnalogBdTemp" Then
                            If Not T.SetPWM("ANALOGBD", 0) Or Not T.SetTEC("ANALOGBD", "HEAT", 0) Then
                                Form1.AppendText(T.ErrorMsg)
                                Return False
                            End If
                        End If
                        'Else
                        '    Form1.AppendText(condition_name + " " + HeaterDT.Rows(row_cnt - 1)(condition_name).ToString + " has not reached " + expected.ToString)
                    End If
                Next
                row_cnt += 1
            ElseIf Now.Subtract(HeaterDT.Rows(row_cnt - 1)("Timestamp")).TotalSeconds > 10 Then
                Form1.AppendText("No heater data in 10s")
                Exit While
            End If
        End While

        T.SetPWM("OILBLOCK", 0)
        T.SetTEC("OILBLOCK", "HEAT", 0)
        T.SetPWM("ANALOGBD", 0)
        T.SetTEC("ANALOGBD", "HEAT", 0)

        If conditions_met Then
            Form1.AppendText("zones reach expected values in " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString)
        Else
            For Each condition_name In conditions.Keys
                If Not conditions(condition_name).met Then
                    Form1.AppendText("condition " + condition_name + " not met")
                End If
            Next
        End If
        Form1.AppendText("-----------------------------------------------------")

        Return conditions_met
    End Function

    Function WaitHeaterCondition(ByVal conditions As Hashtable, ByVal Timeout As Integer) As Boolean
        Dim row_cnt As Integer
        Dim conditions_met As Boolean = False
        Dim startTime As DateTime = Now
        Dim condition_name As String
        Dim limit As Double
        Dim RowDateTime As DateTime
        Dim WhereClause As String
        Dim ts_60s As New TimeSpan(0, 0, 60)
        Dim minimum, maximum, average, expected As Double

        row_cnt = HeaterDT.Rows.Count

        ' wait for pwm conditions
        conditions_met = False
        While (HeaterThread.IsAlive And Not conditions_met And Now.Subtract(startTime).TotalSeconds < Timeout)
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString

                conditions_met = True
                For Each condition_name In conditions.Keys
                    If Not condition_name.EndsWith("PWM") Then Continue For
                    If HeaterDT.Rows(row_cnt)(condition_name) < conditions(condition_name).pwm Then
                        conditions_met = False
                    Else
                        conditions(condition_name).met = True
                    End If
                Next
                row_cnt += 1
            End If
        End While
        If Not conditions_met Then
            For Each condition_name In conditions.Keys
                If condition_name.EndsWith("PWM") Then
                    If Not conditions(condition_name).met Then
                        Form1.AppendText("Timeout waiting for " + condition_name + " to reach " + conditions(condition_name).pwm.ToString)
                    End If
                End If
            Next
            Return False
        End If

        ' wait for temp to be stable for 60s
        conditions_met = False
        While (HeaterThread.IsAlive And Not conditions_met And Now.Subtract(startTime).TotalSeconds < Timeout)
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
                System.Threading.Thread.Sleep(50)
                If HeaterDT.Rows.Count > 60 Then
                    RowDateTime = HeaterDT.Rows(row_cnt)("Timestamp")
                    For Each condition_name In conditions.Keys
                        If Not condition_name.EndsWith("Temp") Or conditions(condition_name).met Then Continue For
                        limit = conditions(condition_name).limit
                        expected = conditions(condition_name).expected
                        WhereClause = "Timestamp > '" + RowDateTime.Subtract(ts_60s).ToString + "'"
                        minimum = HeaterDT.Compute("MIN(" + condition_name + ")", WhereClause)
                        maximum = HeaterDT.Compute("MAX(" + condition_name + ")", WhereClause)
                        'average = HeaterDT.Compute("AVG(" + condition_name + ")", WhereClause)
                        If conditions(condition_name).stop_when_stable Then
                            'Form1.AppendText("Min = " + minimum.ToString + ", Max = " + maximum.ToString + ", Limit" + limit.ToString)
                            'If (maximum - minimum) < limit And ((average > (expected - 2.0 * limit)) And (average < expected + 2.0 * limit)) Then
                            If ((maximum - minimum) < limit) Then
                                conditions(condition_name).met = True
                                Form1.AppendText(condition_name + " has stayed within the limit " + limit.ToString)
                                'Else
                                '    Form1.AppendText(condition_name + " has not stayed within the limit " + limit.ToString)
                            End If
                        Else
                            Dim loLimit As Double = expected - limit
                            Dim hiLimit As Double = expected + limit
                            'Form1.AppendText("Min = " + minimum.ToString + ", Max = " + maximum.ToString + ", LowLimit = " + loLimit.ToString + ", HiLimit = " + hiLimit.ToString)
                            If (minimum > expected - limit) And (maximum < expected + limit) Then
                                conditions(condition_name).met = True
                                Form1.AppendText(condition_name + " has stayed within the range " + loLimit.ToString + " - " + hiLimit.ToString)
                                'Else
                                '    Form1.AppendText(condition_name + " has not stayed within the range " + loLimit.ToString + " - " + hiLimit.ToString)
                            End If
                        End If
                    Next
                End If

                conditions_met = True
                For Each condition_name In conditions.Keys
                    If Not conditions(condition_name).met Then
                        conditions_met = False
                        Exit For
                    End If
                    'If condition_name.EndsWith("Temp") Then
                    '    conditions(condition_name).met = False
                    'End If
                Next
                row_cnt += 1
            End If
        End While

        Return conditions_met
    End Function

    Function WaitH2scanReady(ByVal Timeout As Integer) As Boolean
        Dim H2SCAN_ready As Boolean = False
        Dim startTime As DateTime
        Dim row_cnt As Integer

        startTime = Now
        row_cnt = HeaterDT.Rows.Count
        While (HeaterThread.IsAlive And Not H2SCAN_ready And Now.Subtract(startTime).TotalSeconds < Timeout)
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
                Try
                    If HeaterDT.Rows(row_cnt)("Status") = 8001 Then
                        H2SCAN_ready = True
                    End If
                Catch ex As Exception

                End Try
                row_cnt += 1
            End If
        End While
        If Not H2SCAN_ready Then
            Form1.AppendText("Timeout waiting for H2SCAN to be ready")
        Else
            Form1.AppendText("H2SCAN ready in " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString)
        End If

        Return H2SCAN_ready
    End Function

    Function WaitReworkH2scanReady(ByVal Timeout As Integer) As Boolean
        Dim H2SCAN_ready As Boolean = False
        Dim startTime As DateTime
        Dim row_cnt As Integer

        startTime = Now
        row_cnt = HeaterDT.Rows.Count
        While (HeaterThread.IsAlive And Not H2SCAN_ready And Now.Subtract(startTime).TotalSeconds < Timeout)
            Application.DoEvents()

            ' Check for stop flag
            If Stopped Then
                Form1.AppendText("Stop button was pressed.")
                Exit While
            End If

            If HeaterDT.Rows.Count > row_cnt Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
                Try
                    If HeaterDT.Rows(row_cnt)("Status") <> 4096 Then
                        H2SCAN_ready = True
                    End If
                Catch ex As Exception

                End Try
                row_cnt += 1
            End If
        End While
        If Not H2SCAN_ready Then
            Form1.AppendText("Timeout waiting for H2SCAN to be ready")
        Else
            Form1.AppendText("H2SCAN ready in " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString)
        End If

        Return H2SCAN_ready
    End Function
End Class

Public Class condition
    Private _expected As Double
    Private _limit As Double
    Private _check As String
    Private _met As Boolean = False
    Private _stop_when_stable As Boolean = False
    Private _pwm As Double

    Public Property expected() As Double
        Set(value As Double)
            _expected = value
        End Set
        Get
            Return _expected
        End Get
    End Property

    Public Property limit() As Double
        Set(value As Double)
            _limit = value
        End Set
        Get
            Return _limit
        End Get
    End Property

    Public Property check() As String
        Set(value As String)
            _check = value
        End Set
        Get
            Return _check
        End Get
    End Property

    Public Property met() As Boolean
        Set(value As Boolean)
            _met = value
        End Set
        Get
            Return _met
        End Get
    End Property

    Public Property pwm() As Double
        Get
            Return _pwm
        End Get
        Set(value As Double)
            _pwm = value
        End Set
    End Property

    Public Property stop_when_stable() As Boolean
        Set(value As Boolean)
            _stop_when_stable = value
        End Set
        Get
            Return _stop_when_stable
        End Get
    End Property
End Class
'