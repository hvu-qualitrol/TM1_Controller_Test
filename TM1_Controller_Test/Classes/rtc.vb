Imports System.Text.RegularExpressions

Public Class rtc
    Private _ErrorMsg As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Function Clear() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        If Not SF.Cmd(SerialPort, Response, "rtc -z", 5) Then
            _ErrorMsg = "Problem sending cmd 'rtc -z'"
            Return False
        End If

        Return True
    End Function

    Function SetCapacitance(ByVal capacitance As Integer)
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim rtc_info As Hashtable

        Command = "rtc -s " + capacitance.ToString

        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "Problem sending cmd " + Command
            Return False
        End If

        If Not GetRTC(rtc_info) Then
            Return False
        End If

        If Not rtc_info("capacitance") = capacitance Then
            _ErrorMsg = "Problem setting RTC Capacitance." + vbCr
            _ErrorMsg += "Trying to set to " + capacitance.ToString
            If rtc_info.Contains("capacitance") Then
                _ErrorMsg += "Actual = " + rtc_info("capacitance").ToString
            End If
            Return False
        End If

        Return True
    End Function

    Function SetOffset(ByVal frequency As Double) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim Temperature As Double = 0
        Dim Offset As Double = 1000
        Dim Line As String

        Command = "rtc -f " + frequency.ToString

        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "Problem sending cmd " + Command
            Return False
        End If

        For Each Line In Response.Split(Chr(13))
            Form1.AppendText(Line)
            If Regex.IsMatch(Line, "freq\s+\d+\.\d+\s+at temp\s+\d+\.\d+") Then
                Temperature = CDbl(Regex.Split(Line, "freq\s+\d+\.\d+\s+at temp\s+(\d+\.\d+)")(1))
            End If
            If Regex.IsMatch(Line, "offset\s+.*\d+\.\d+\s+ppm") Then
                Offset = CDbl(Regex.Split(Line, "offset\s+(.*\d+\.\d+)\s+ppm")(1))
            End If
        Next

        If Temperature < 23 Or Temperature > 30 Then
            '_ErrorMsg = "Expected temperature between 23C and 30C"
            _ErrorMsg = "Expected temperature between 27C and 34C"
            Return False
        End If

        If Math.Abs(Offset) > 20 Then
            _ErrorMsg = "Expected offset < +/- 20ppm"
            Return False
        End If

        Return True
    End Function

    Function GetRTC(ByRef rtc_settings As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim load_capacitance As Integer

        rtc_settings = New Hashtable

        If Not SF.Cmd(SerialPort, Response, "rtc", 5, "> ", True) Then
            _ErrorMsg = "Problem sending cmd 'rtc'"
            Return False
        End If

        For Each Line In Response.Split(Chr(13))
            Line = Regex.Replace(Line, Chr(10), "")
            Line = Regex.Replace(Line, "^\s+", "")
            Line = Regex.Replace(Line, "\s+:\s+", ":")
            Line = Regex.Replace(Line, "\s+$", "")
            Fields = Split(Line, ":")
            If Fields.Length = 2 Then
                If Fields(1).Contains("ppm") Then
                    Fields(1) = Regex.Replace(Fields(1), "ppm", "")
                End If
                rtc_settings.Add(Fields(0), Fields(1))

                If Fields(0) = "RTC_CR" Then
                    load_capacitance = CInt(Regex.Split(Fields(1), "\((\d+)pf\)")(1))
                    rtc_settings.Add("capacitance", load_capacitance)
                End If
            End If
        Next

        Return True
    End Function
End Class
