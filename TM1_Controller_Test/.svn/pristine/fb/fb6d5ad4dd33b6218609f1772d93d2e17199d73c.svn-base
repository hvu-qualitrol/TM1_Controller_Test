Imports System.IO.Ports
Imports System.Text.RegularExpressions

'Public Delegate Sub CallBackDelegate(ByVal status As String)
Public Delegate Sub CallBackDelegate(ByVal row As DataRow)

Public Class Leak_Fixture
    ' common variables with background thread
    Friend SerialPortName As String
    Friend PressureDT As DataTable
    Friend ErrorMsg As String = "UNKNOWN"
    Friend StopLogging As Boolean = False
    Friend RetVal As Boolean
    Friend SetValves As Boolean = False
    Friend NewValveSetting As Integer = 0

    Public LeakThread As System.Threading.Thread

    Private m_BaseControl As Control
    Private m_CallBackFunction As CallBackDelegate

    Sub New(ByRef caller As Control, ByRef callbackFunction As CallBackDelegate)
        m_BaseControl = caller
        m_CallBackFunction = callbackFunction
    End Sub

    Sub StartLogging()
        LeakThread = New System.Threading.Thread(AddressOf LogPressure)
        LeakThread.Start()
    End Sub

    Sub LogPressure()
        Dim LeakFixtureSP As New SerialPort(SerialPortName, 2400, 0, 8, 1)
        Dim Line As String
        Dim Fields() As String
        Dim PressureRaw As Integer
        Dim ValveSettings As Integer = 0
        Dim row As DataRow
        Dim startTime As DateTime = Now

        Try
            LeakFixtureSP.Open()
        Catch ex As Exception
            ErrorMsg = "Problem opening leak fixture serial port"
            RetVal = False
            Exit Sub
        End Try
        LeakFixtureSP.DiscardInBuffer()
        LeakFixtureSP.ReadTimeout = 2000
        LeakFixtureSP.NewLine = Chr(10)

        While (Not StopLogging)
            'Application.DoEvents()
            If Stopped Then
                ErrorMsg = "LogPressure(): Test stopped."
                'StopLogging = True
                Exit While
            End If
            Try
                Line = LeakFixtureSP.ReadLine()
                Line = Regex.Replace(Line, Chr(13), "")
                If Regex.IsMatch(Line, "PRESSURE=\d+,\s+B0=\d+$") Then
                    Fields = Regex.Split(Line, "PRESSURE=(\d+),\s+B0=(\d+)")
                    PressureRaw = CInt(Fields(1))
                    ValveSettings = CInt(Fields(2))
                    If SetValves Then
                        If ValveSettings = NewValveSetting Then
                            SetValves = False
                        Else
                            LeakFixtureSP.Write(":" + Chr(NewValveSetting))
                        End If
                    End If
                    row = PressureDT.NewRow()
                    row("Timestamp") = Now
                    row("TimeSinceStart") = Math.Round(Now.Subtract(startTime).TotalSeconds, 1)
                    row("ValveSettings") = ValveSettings
                    row("Pressure") = Math.Round(PressureRaw * 100.0 / 1023, 2)
                    m_BaseControl.Invoke(m_CallBackFunction, row)
                    'PressureDT.Rows.Add(row)
                End If
            Catch ex As Exception

            End Try
        End While
        NewValveSetting = 12
        LeakFixtureSP.Write(":" + Chr(NewValveSetting))
        ErrorMsg = "Setting valve to 12"
        System.Threading.Thread.Sleep(300)
        LeakFixtureSP.Close()
        'System.Threading.Thread.Sleep(300)

    End Sub

    Sub LogPressure0()
        Dim LeakFixtureSP As New SerialPort(SerialPortName, 2400, 0, 8, 1)
        Dim Line As String
        Dim Fields() As String
        Dim PressureRaw As Integer
        Dim ValveSettings As Integer = 0
        Dim row As DataRow
        Dim startTime As DateTime = Now

        Try
            LeakFixtureSP.Open()
        Catch ex As Exception
            ErrorMsg = "Problem opening leak fixture serial port"
            RetVal = False
            Exit Sub
        End Try
        LeakFixtureSP.DiscardInBuffer()
        LeakFixtureSP.ReadTimeout = 2000
        LeakFixtureSP.NewLine = Chr(10)

        While (Not StopLogging)
            Application.DoEvents()
            If Stopped Then
                ErrorMsg = "LogPressure(): Test stopped."
                'StopLogging = True
                Exit While
            End If
            Try
                Line = LeakFixtureSP.ReadLine()
                Line = Regex.Replace(Line, Chr(13), "")
                If Regex.IsMatch(Line, "PRESSURE=\d+,\s+B0=\d+$") Then
                    Fields = Regex.Split(Line, "PRESSURE=(\d+),\s+B0=(\d+)")
                    PressureRaw = CInt(Fields(1))
                    ValveSettings = CInt(Fields(2))
                    If SetValves Then
                        If ValveSettings = NewValveSetting Then
                            SetValves = False
                        Else
                            LeakFixtureSP.Write(":" + Chr(NewValveSetting))
                        End If
                    End If
                    row = PressureDT.NewRow()
                    row("Timestamp") = Now
                    row("TimeSinceStart") = Math.Round(Now.Subtract(startTime).TotalSeconds, 1)
                    row("ValveSettings") = ValveSettings
                    row("Pressure") = Math.Round(PressureRaw * 100.0 / 1023, 2)
                    m_BaseControl.Invoke(m_CallBackFunction, row)
                    'PressureDT.Rows.Add(row)
                End If
            Catch ex As Exception

            End Try
        End While
        NewValveSetting = 12
        LeakFixtureSP.Write(":" + Chr(NewValveSetting))
        ErrorMsg = "Setting valve to 12"
        System.Threading.Thread.Sleep(300)
        LeakFixtureSP.Close()
        'System.Threading.Thread.Sleep(300)

    End Sub

    Sub NewPressureTable()

        PressureDT = New DataTable()
        PressureDT.Columns.Add("Timestamp", Type.GetType("System.DateTime"))
        PressureDT.Columns.Add("TimeSinceStart", Type.GetType("System.Double"))
        PressureDT.Columns.Add("ValveSettings", Type.GetType("System.Int32"))
        PressureDT.Columns.Add("Pressure", Type.GetType("System.Double"))
        'AddHandler PressureDT.RowChanged, New DataRowChangeEventHandler(AddressOf NewPressureData)

    End Sub

    'Sub NewPressureData(ByVal sender As Object, ByVal e As DataRowChangeEventArgs)
    '    m_BaseControl.Invoke(m_CallBackFunction, "HERE")
    'End Sub

    Function ChangeValves(ByVal valve_setting As Integer)
        If Stopped Then
            ErrorMsg = "ChangeValves(): Test aborted."
            Return False
        End If
        Dim startTime As DateTime = Now
        NewValveSetting = valve_setting
        SetValves = True

        While (SetValves And Now.Subtract(startTime).TotalSeconds < 10)
            Application.DoEvents()
            If Stopped Then
                ErrorMsg = "ChangeValves(): Test aborted."
                Return False
            End If
        End While
        If SetValves Then
            ErrorMsg = "Timeout changing leak fixture valve setting to " + valve_setting.ToString
            Return False
            SetValves = False
            Return False
        End If

        Return True
    End Function

    Function CalculateLeakRate2(ByVal start As Double, ByVal finish As Double) As Double
        Dim slope As Double = 0
        Dim sum_square_x_minus_xave As Double = 0
        Dim sum_x_minus_xave_times_y_minus_yave As Double = 0
        Dim WhereClause As String
        Dim avgP1 As Double = 0
        Dim avgP2 As Double = 0

        'Calculate the average pressure of the first minite
        WhereClause = "TimeSinceStart >= " + start.ToString + " and " +
            "TimeSinceStart < " + (start + 60.0).ToString
        avgP1 = PressureDT.Compute("AVG(Pressure)", WhereClause)

        'Calculate the average pressure of the last minite
        WhereClause = "TimeSinceStart >= " + (finish - 60.0).ToString + " and " +
            "TimeSinceStart < " + finish.ToString
        avgP2 = PressureDT.Compute("AVG(Pressure)", WhereClause)

        ' Calculate the leak rate in PSI/second
        Dim leakRate As Double
        leakRate = Math.Round((60.0 * (avgP2 - avgP1) / (finish - start)), 4)

        Return leakRate
    End Function

    Function CalculatePressureStdev(ByVal start As Double, ByVal finish As Double) As Double
        Dim WhereClause As String
        Dim stdev As Double

        'Calculate the standard deviation
        WhereClause = "TimeSinceStart >= " + start.ToString + " and " +
            "TimeSinceStart < " + finish.ToString
        stdev = PressureDT.Compute("STDEV(Pressure)", WhereClause)

        Return stdev
    End Function

    Function CalculateLeakRate1(ByVal start As Double, ByVal finish As Double) As Double
        Dim slope As Double = 0

        'Calculate the average pressures of the period of one minute at the two ends
        Dim avgP1 As Double = 0
        Dim avgP2 As Double = 0
        Dim rowCount1 As Integer = 0
        Dim rowCount2 As Integer = 0
        For i = 0 To PressureDT.Rows.Count - 1
            If PressureDT.Rows(i)("TimeSinceStart") >= start And PressureDT.Rows(i)("TimeSinceStart") < (start + 60.0) Then
                avgP1 += PressureDT.Rows(i)("Pressure")
                rowCount1 += 1
            End If
            If PressureDT.Rows(i)("TimeSinceStart") >= (finish - 60.0) And PressureDT.Rows(i)("TimeSinceStart") < finish Then
                avgP2 += PressureDT.Rows(i)("Pressure")
                rowCount2 += 1
            End If
        Next
        avgP1 /= rowCount1
        avgP2 /= rowCount2

        ' Calculate the leak rate in PSI/second
        Dim leakRate As Double
        leakRate = Math.Round((60.0 * (avgP2 - avgP1) / (finish - start)), 4)

        Return leakRate
    End Function

    Function CalculateLeakRate(ByVal start As Double, ByVal finish As Double) As Double
        Dim sum_x As Double
        Dim sum_y As Double
        Dim xave As Double
        Dim yave As Double
        Dim x As Double
        Dim y As Double
        Dim slope As Double = 0
        Dim sum_square_x_minus_xave As Double = 0
        Dim sum_x_minus_xave_times_y_minus_yave As Double = 0
        Dim WhereClause As String
        Dim row1, row2 As Integer

        WhereClause = "TimeSinceStart >= " + start.ToString + " and " +
            "TimeSinceStart <= " + finish.ToString
        sum_x = PressureDT.Compute("SUM(TimeSinceStart)", WhereClause)
        sum_y = PressureDT.Compute("SUM(Pressure)", WhereClause)
        For i = 0 To PressureDT.Rows.Count - 1
            If PressureDT.Rows(i)("TimeSinceStart") = start Then
                row1 = i
            End If
            If PressureDT.Rows(i)("TimeSinceStart") = finish Then
                row2 = i
            End If
        Next
        xave = sum_x / (row2 - row1 + 1)
        yave = sum_y / (row2 - row1 + 1)

        For i = 0 To PressureDT.Rows.Count - 1
            If PressureDT.Rows(i)("TimeSinceStart") >= start And PressureDT.Rows(i)("TimeSinceStart") <= finish Then
                x = PressureDT.Rows(i)("TimeSinceStart")
                sum_square_x_minus_xave = sum_square_x_minus_xave + (x - xave) * (x - xave)
                y = PressureDT.Rows(i)("Pressure")
                sum_x_minus_xave_times_y_minus_yave += (x - xave) * (y - yave)
                slope = Math.Round(sum_x_minus_xave_times_y_minus_yave / sum_square_x_minus_xave, 4)
            End If
        Next

        Return slope
    End Function
End Class
