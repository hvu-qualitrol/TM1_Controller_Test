Imports System.Text.RegularExpressions
Imports TM1_Controller_Test.NTPServer

Public Class TM1
    'TODO add errmessage return
    Dim RelayConnections As Hashtable
    Private _ErrorMsg As String = ""

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property


    Public Sub New()
        Dim R As Relay

        RelayConnections = New Hashtable
        _ErrorMsg = ""

        R = New Relay
        R.COM = 0
        R.NC = 8
        R.NO = 9
        RelayConnections.Add("Relay1", R)

        R = New Relay
        R.COM = 1
        R.NC = 10
        R.NO = 11
        RelayConnections.Add("Relay2", R)

        R = New Relay
        R.COM = 2
        R.NC = 12
        R.NO = 13
        RelayConnections.Add("Relay3", R)

        R = New Relay
        R.COM = 3
        R.NC = 14
        R.NO = 15
        RelayConnections.Add("Relay4", R)

    End Sub

    Public Function GetRelayConnection(ByVal Relay As Integer, ByVal Terminal As String) As Integer
        Dim RelayName As String = "Relay" + Relay.ToString

        If Terminal = "COM" Then
            Return RelayConnections(RelayName).COM
        End If

        If Terminal = "NO" Then
            Return RelayConnections(RelayName).NO
        End If

        If Terminal = "NC" Then
            Return RelayConnections(RelayName).NC
        End If

    End Function


    Public Function SetRelay(ByVal SerialPort As Object, ByVal Relay As Integer, ByVal State As RelayState) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim RelayStateString As String
        Dim RelayStatus As Hashtable
        Dim RelayName As String = "Relay" + Relay.ToString
        Dim retryCnt As Integer = 0
        Dim success As Boolean = False

        Select Case State
            Case RelayState.RELAY_OFF
                RelayStateString = "OFF"

            Case RelayState.RELAY_ON
                RelayStateString = "ON"

            Case Else
                Return False

        End Select

        While Not success And retryCnt < 3
            If retryCnt > 1 Then
                System.Threading.Thread.Sleep(300)
                Form1.AppendText("retrying relay set")
            End If
            retryCnt += 1
            If Not SF.Cmd(SerialPort, Response, "relay " + Relay.ToString + " " + RelayStateString, 5) Then
                Continue While
                'Return False
            End If

            System.Threading.Thread.Sleep(50)
            If Not GetRelays(SerialPort, RelayStatus) Then
                Continue While
                'Return False
            End If

            If Not RelayStatus(RelayName) = RelayStateString Then
                _ErrorMsg = "RelayStatus = " + RelayStatus(RelayName).ToString + ", expected " + RelayStateString
                Continue While
                'Return False
            Else
                success = True
            End If
        End While

        Return success
        'Return True
    End Function

    Public Function GetRelays(ByVal SerialPort As Object, ByRef RelayStatus As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim relay As Integer
        Dim retryCnt As Integer = 0
        Dim success As Boolean = False

        While Not success And retryCnt < 3
            success = True
            retryCnt += 1

            RelayStatus = New Hashtable
            ' System.Threading.Thread.Sleep(50)
            If Not SF.Cmd(SerialPort, Response, "relay", 10) Then
                Return False
            End If
            For Each Line In Response.Split(Chr(13))
                Line = Line.Trim
                If Line.StartsWith("Relay") Then
                    Fields = Regex.Split(Line, "\s+")
                    Try
                        RelayStatus.Add("Relay" + Fields(1), Fields(2))
                    Catch ex As Exception
                        _ErrorMsg = ex.ToString
                        Return False
                    End Try
                End If
            Next
            For relay = 1 To 4
                If Not RelayStatus.Contains("Relay" + relay.ToString) Then
                    success = False
                    RelayStatus.Add("Relay" + relay.ToString, "UKNOWN")
                End If
            Next
        End While

        Return success
    End Function

    Function GetPumpRpm() As Hashtable
        Dim SF As New SerialFunctions
        Dim PumpCondition As New Hashtable
        Dim Response As String
        Dim Fields() As String

        PumpCondition.Add("STATUS", False)
        PumpCondition.Add("ERROR_MESSAGE", "")
        If Not SF.Cmd(SerialPort, Response, "pump", 10) Then
            Return PumpCondition
        End If

        For Each Line In Response.Split(Chr(13), Chr(10))
            'If Line.StartsWith("Oil Pump:") Then
            If Regex.IsMatch(Line, "Oil Pump:\s+(\w+)\s+\((\d+\.\d+)\s+rps\)") Then
                Fields = Regex.Split(Line, "Oil Pump:\s+(\w+)\s+\((\d+\.\d+)\s+rps\)")
                PumpCondition.Add("DIR", Fields(1))
                Try
                    PumpCondition.Add("RPM", CDbl(Fields(2)))
                Catch ex As Exception
                    PumpCondition("ERROR_MESSAGE") = "Problem converting " + Fields(2) + " to rpm"
                    Return PumpCondition
                End Try
                PumpCondition("STATUS") = True
            End If
        Next

        Return PumpCondition
    End Function

    Function SetPumpSpeed(ByVal Dir As String, ByVal Percentage As Integer) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String

        Command = "pump " + Dir + " " + Percentage.ToString
        If Not SF.Cmd(SerialPort, Response, "pump " + Dir + " " + Percentage.ToString, 10) Then
            _ErrorMsg = "Problem sending command " + Command
            Return False
        End If
        If Response.Contains("Usage") Or Response.Contains("Error") Then
            _ErrorMsg = Response
            Return False
        End If

        Return True
    End Function

    Function PumpOff() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(SerialPort, Response, "pump off", 10) Then
            _ErrorMsg = "Problem sending command 'pump off'"
            Return False
        End If
        If Response.Contains("Usage") Or Response.Contains("Error") Then
            _ErrorMsg = Response
            Return False
        End If

        Return True
    End Function

    Function GetVersionInfo(ByRef VersionInfo As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Fields() As String

        VersionInfo = New Hashtable
        If Not SF.Cmd(SerialPort, Response, "ver", 5) Then
            Form1.AppendText("failed sending cmd ver", True)
            Return False
        End If
        'Form1.AppendText(Response, True)
        For Each Line In Response.Split(Chr(13))
            Line = Regex.Replace(Line, Chr(10), "")
            Form1.AppendText(Line, True)
            Line = Regex.Replace(Line, "^\s+", "")
            Line = Regex.Replace(Line, "\s+=\s+", "=")
            Line = Regex.Replace(Line, "\s+$", "")
            Fields = Split(Line, "=")
            If Fields.Length = 2 Then
                VersionInfo.Add(Fields(0), Fields(1))
            End If
        Next

        Return True
    End Function

    Function SetConfig(ByVal ConfigItems As Hashtable, ByRef NeedsReboot As Boolean, Optional ByVal Reboot As Boolean = False) As Boolean
        Dim CurrentConfig As Hashtable
        Dim ConfigName As String
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim ConfigSetSuccessfull As Boolean
        Dim results As ReturnResults

        If Not GetConfig(CurrentConfig) Then
            Return False
        End If

        NeedsReboot = False
        For Each ConfigName In ConfigItems.Keys
            If Not CurrentConfig(ConfigName) = ConfigItems(ConfigName) Then
                Command = "config -s set " + ConfigName + " " + ConfigItems(ConfigName)
                If Not SF.Cmd(SerialPort, Response, Command, 5) Then
                    Form1.AppendText("Problem sending command " + Command)
                    Return False
                End If
                NeedsReboot = True
            End If
        Next

        If NeedsReboot Then
            CommonLib.Delay(10)
        End If
        If Not GetConfig(CurrentConfig) Then
            Return False
        End If
        ConfigSetSuccessfull = True
        _ErrorMsg = ""
        For Each ConfigName In ConfigItems.Keys
            If Not CurrentConfig(ConfigName) = ConfigItems(ConfigName) Then
                ConfigSetSuccessfull = False
                _ErrorMsg += ConfigName + "=" + CurrentConfig(ConfigName) + ", Expected=" + ConfigItems(ConfigName) + vbCr
            End If
        Next

        If ConfigSetSuccessfull And NeedsReboot And Reboot Then
            results = SF.Reboot(SerialPort)
            If Not results.PassFail Then
                Form1.AppendText("Problem rebooting")
                Form1.AppendText(SF.ErrorMsg)
                Return False
            End If
            If Not GetConfig(CurrentConfig) Then
                Return False
            End If
            ConfigSetSuccessfull = True
            _ErrorMsg = ""
            For Each ConfigName In ConfigItems.Keys
                If Not CurrentConfig(ConfigName) = ConfigItems(ConfigName) Then
                    ConfigSetSuccessfull = False
                    _ErrorMsg += "ConfigName=" + CurrentConfig(ConfigName) + ", Expected=" + ConfigItems(ConfigName) + vbCr
                End If
            Next
        End If

        Return ConfigSetSuccessfull
    End Function

    Function GetConfig(ByRef Config As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Fields() As String
        Dim Name, Value As String

        Config = New Hashtable
        'If Not SF.Cmd(SerialPort, Response, "config", 5, "> ", True, True, False) Then
        If Not SF.Cmd(SerialPort, Response, "config", 5, "> ", True, False, False) Then
            _ErrorMsg = "failed sending cmd config"
            Return False
        End If

        'System.Text.Encoding utf_8 = System.Text.Encoding.UTF8;

        ' // This is our Unicode string:
        ' string s_unicode = "abcéabc";

        ' // Convert a string to utf-8 bytes.
        ' byte[] utf8Bytes = System.Text.Encoding.UTF8.GetBytes(s_unicode);

        ' // Convert utf-8 bytes to a string.
        ' string s_unicode2 = System.Text.Encoding.UTF8.GetString(utf8Bytes);

        '    MessageBox.Show(s_unicode2);

        For Each Line In Response.Split(Chr(13))
            Line = Regex.Replace(Line, Chr(10), "")
            Fields = Regex.Split(Line, "=")
            Try
                Name = Fields(0)
                Value = Fields(1)
            Catch ex As Exception
                _ErrorMsg = "Problem extracting name/value from " + Line
                Return False
            End Try
            Name = Regex.Replace(Name, "^\s+", "")
            Name = Regex.Replace(Name, "\s+$", "")
            Value = Regex.Replace(Value, "^\s+", "")
            Value = Regex.Replace(Value, "\s+$", "")
            Config.Add(Name, Value)
        Next

        Return True
    End Function

    Function SetLED(ByVal TargetLED As Integer, ByVal LedState As Integer, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim T As New TM1
        Dim SF As New SerialFunctions
        Dim cmd As String
        Dim Response As String

        If Not TargetLED = LED.ALARM And Not TargetLED = LED.POWER And Not TargetLED = LED.SERVICE Then
            Form1.AppendText(TargetLED.ToString + " is not a valid LED", True)
            Return False
        End If

        If Not LedState = State._ON And Not LedState = State._OFF Then
            Form1.AppendText("LED state = " + LedState.ToString + " not valid for " + LED_Strings(TargetLED), True)
            Return False
        End If

        cmd = "led " + LED_Strings(TargetLED) + " " + State_Strings(LedState)
        If Not SF.Cmd(SerialPort, Response, cmd, 5, "> ", Quiet) Then
            Form1.AppendText("failed sending cmd ver", True)
            Return False
        End If

        Return True
    End Function

    Function GetUutDate(ByRef ZuluTime As String) As Boolean
        Dim SF As New SerialFunctions

        If Not SF.Cmd(SerialPort, ZuluTime, "date", 5) Then
            ZuluTime = "failed sending cmd date"
            Return False
        End If

        Return True
    End Function

    Function GetRtcCalDateTime(ByRef CalcDateTime As DateTime) As Boolean
        Dim State As Boolean = False
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim DateTimeStr As String

        If Not SF.Cmd(SerialPort, Response, "ev 23 3", 10, "> ", True) Then
            _ErrorMsg = "Problem sending cmd 'ev 23'"
            Return False
        End If

        For Each Line In Response.Split(Chr(13))
            If Line.Contains("SYSTEM MODE CHANGE: RTC freq") Then
                'resultTxtBox.AppendText(Line & Environment.NewLine)
                DateTimeStr = Line.Substring(11, 20)
                'resultTxtBox.AppendText(datetimeStr & Environment.NewLine)
                CalcDateTime = DateTime.Parse(DateTimeStr)
                State = True
                Exit For
            End If
        Next

        Return State

    End Function

    Function GetRtcAllowedDiff() As Double
        ' Calculate the limit based on the RTC calibration date
        Dim RtcCalDateTime As DateTime
        Dim ServerTimeUTC_Str As String = ""
        Dim ServerTimeUTC As DateTime
        Dim DeltaSeconds As Double
        Dim AllowedDiff As Double = -111.0
        Dim RtcDriftPerSecSpec As Double = 10.0 * 60.0 / (365.0 * 24.0 * 3600.0)

        ' Gets the time from the computer.
        If Not CommonLib.GetLocalTime(ServerTimeUTC_Str, "UTC") Then
            Form1.AppendText("Problem getting network time", True)
            Form1.AppendText(ServerTimeUTC_Str, True)
            Return AllowedDiff
        End If
        ServerTimeUTC = ServerTimeUTC_Str

        If GetRtcCalDateTime(RtcCalDateTime) Then
            DeltaSeconds = ServerTimeUTC.Subtract(RtcCalDateTime).TotalSeconds
            AllowedDiff = RtcDriftPerSecSpec * DeltaSeconds
            Form1.AppendText("RTC cal date time: " + RtcCalDateTime)
            Form1.AppendText("Calculated AllowDiff: " + AllowedDiff.ToString)
        End If

        Return AllowedDiff

    End Function

    Function VerifyDate() As Boolean
        Dim ServerTimeUTC_Str As String = ""
        Dim ServerTimeUTC As DateTime
        Dim ZuluTime As String = ""
        Dim UutTimeUTC As DateTime
        Dim AllowedDiff As Integer

        If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Or Form1.AssemblyLevel_ComboBox.Text = "W223_REWORK" Then
            AllowedDiff = 5
        Else
            AllowedDiff = 1
        End If

        If Not GetUutDate(ZuluTime) Then
            Form1.AppendText(ZuluTime)
            Return False
        End If
        Form1.AppendText("UUT Time (ZULU) = " + ZuluTime)

        ' Gets the time from the computer.
        If Not CommonLib.GetLocalTime(ServerTimeUTC_Str, "UTC") Then
            Form1.AppendText("Problem getting network time", True)
            Form1.AppendText(ServerTimeUTC_Str, True)
            Return False
        End If

        'ServerTimeUTC_Str = GetNetworkTime().ToString()
        'Form1.AppendText("NTP Time (UTC) = " + ServerTimeUTC_Str, True)

        ServerTimeUTC = ServerTimeUTC_Str
        Form1.AppendText("Server Time (UTC) = " + ServerTimeUTC)


        If Not CommonLib.ZuluToUTC(ZuluTime, UutTimeUTC) Then
            Form1.AppendText("Problem converting ZuluTime to DateTime type")
            Return False
        End If
        Form1.AppendText("UUT Time (UTC) = " + UutTimeUTC)
        Form1.AppendText("DiffTime = " + ServerTimeUTC.Subtract(UutTimeUTC).TotalSeconds.ToString + " seconds")

        'Calculate the limit based on the spec & the delta time between the cal date and now
        Dim CalcLimit As Double = GetRtcAllowedDiff()
        If AllowedDiff < CalcLimit Then
            AllowedDiff = CalcLimit
        End If

        'Check to see that the PC time and the monitor are within the allowed time.
        If Math.Abs(ServerTimeUTC.Subtract(UutTimeUTC).TotalSeconds) > AllowedDiff Then
            Form1.AppendText("Expected diff between server and uut time <= " + AllowedDiff.ToString + " seconds")

            If ZuluTime.Contains("1970") Then
                Form1.AppendText("***************************************************************************")
                Form1.AppendText("Check the installed battery. It could be installed incorrectly or bad.")
                Form1.AppendText("***************************************************************************")
            End If

            Return False
        End If

        Return True
    End Function

    Function SetTEC_AUTO() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim TecStatus As Hashtable

        Command = "TEC AUTO"
        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "failed sending cmd " + Command
            Return False
        End If

        If Not GetTecStatus(TecStatus) Then
            _ErrorMsg = "Problem getting TEC status " + vbCr + _ErrorMsg
            Return False
        End If
        For Each k In TecStatus.Keys
            Form1.AppendText(k + " " + TecStatus(k)("DIR") + " " + TecStatus(k)("RATE").ToString + " " + TecStatus(k)("MODE"))
        Next

        If Not TecStatus("ANALOGBD")("MODE") = "auto" Then
            _ErrorMsg = "Expected AnalogBd mode=auto"
            Return False
        End If

        If Not TecStatus("OILBLOCK")("MODE") = "auto" Then
            _ErrorMsg = "Expected OilBlock mode=auto"
            Return False
        End If

        Return True
    End Function

    Function SetTEC(ByVal tec As String, ByVal direction As String, ByVal setting As Double) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim TecStatus As Hashtable

        If Not tec = "OILBLOCK" And Not tec = "ANALOGBD" Then
            _ErrorMsg = "Code Error:  SetTEC() tec=" + tec + ", expected tec OILBLOCK or ANALOGBD"
            Return False
        End If

        If Not direction = "HEAT" And Not direction = "COOL" Then
            _ErrorMsg = "Code Error:  SetTEC() direction=" + direction + ", expected direction HEAT or COOL"
            Return False
        End If

        If setting < 0 Or setting > 100 Then
            _ErrorMsg = "Code Error:  SetTEC() setting=" + setting.ToString + ", expected 0 - 100.0"
            Return False
        End If

        Command = "TEC " + tec + " " + direction + " " + setting.ToString
        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "failed sending cmd " + Command
            Return False
        End If

        If Not GetTecStatus(TecStatus) Then
            _ErrorMsg = "Problem getting TEC status " + vbCr + _ErrorMsg
            Return False
        End If
        For Each k In TecStatus.Keys
            Form1.AppendText(k + " " + TecStatus(k)("DIR") + " " + TecStatus(k)("RATE").ToString + " " + TecStatus(k)("MODE"))
        Next
        If Not TecStatus(tec)("MODE") = "manual" Then
            _ErrorMsg = tec + " " + TecStatus(tec)("DIR") + " " + TecStatus(tec)("RATE").ToString + " " + TecStatus(tec)("MODE") + vbCr +
                "Expected (manual)"
            Return False
        End If
        If Not TecStatus(tec)("DIR") = direction Then
            _ErrorMsg = tec + " " + TecStatus(tec)("DIR") + " " + TecStatus(tec)("RATE").ToString + " " + TecStatus(tec)("MODE") + vbCr +
                "Expected " + direction
            Return False
        End If
        If Not TecStatus(tec)("RATE") = setting Then
            _ErrorMsg = tec + " " + TecStatus(tec)("DIR") + " " + TecStatus(tec)("RATE").ToString + " " + TecStatus(tec)("MODE") + vbCr +
                "Expected " + setting.ToString
            Return False
        End If

        Return True
    End Function

    Function GetTecStatus(ByRef TecStatus As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim Status As Hashtable
        Dim retryCnt As Integer = 0
        Dim success As Boolean = False
        Dim tec As String
        Dim param As String

        While Not success And retryCnt < 3
            _ErrorMsg = ""
            If retryCnt > 0 Then
                CommonLib.Delay(2)
            End If
            retryCnt += 1

            System.Threading.Thread.Sleep(50)
            If Not SF.Cmd(SerialPort, Response, "tec status", 10) Then
                _ErrorMsg = "failed sending cmd config"
                Return False
            End If
            TecStatus = New Hashtable
            For Each Line In Response.Split(Chr(13), Chr(10))
                If Regex.IsMatch(Line, "(\w+)\s+(\w+)\s+(\d+\.\d+)\s+\((\w+)\)") Then
                    Fields = Regex.Split(Line, "(\w+)\s+(\w+)\s+(\d+\.\d+)\s+\((\w+)\)")
                    Status = New Hashtable
                    Status.Add("DIR", Fields(2))
                    Try
                        Status.Add("RATE", CDbl(Fields(3)))
                    Catch ex As Exception
                        _ErrorMsg = "Could not extract tec rate from " + Line
                        Return False
                    End Try
                    Status.Add("MODE", Fields(4))
                    TecStatus.Add(Fields(1).ToUpper, Status)
                End If
            Next
            success = True
            For Each tec In {"OILBLOCK", "ANALOGBD"}
                If Not TecStatus.Contains(tec) Then
                    _ErrorMsg = "no tec status for " + tec
                    success = False
                    Exit For
                End If
                For Each param In {"DIR", "RATE", "MODE"}
                    If Not TecStatus(tec).contains(param) Then
                        _ErrorMsg += tec + " tec status missing " + param + vbCr
                    End If
                Next
                If Not _ErrorMsg = "" Then
                    success = False
                    Exit For
                End If
            Next
        End While

        Return success
    End Function

    Function Heater(ByRef Results As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String = ""
        Dim DataLine As String = ""
        Dim Fields() As String
        Dim DataLineFields() As String = {"", "H2_OIL|I", "H2_DGA.PPM|I", "H2.PPM|I", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "Status|I",
        "HCurrent|I", "ResAdc|I", "AdjRes|I", "H2Res.PPM|I", "OilBlockTemp|D", "OilBlockPWM|D", "OilBlockDir|S", "AnalogBdTemp|D", "AnalogBdPWM|D",
        "AnalogBdDir|S", "AmbientTemp|D", "OilblockRevPWM|D", "AnalogBdRevPWM|D", "Tach|I"}
        Dim Name, Type As String
        Dim i As Integer

        System.Threading.Thread.Sleep(150)
        If Not SF.Cmd(SerialPort, Response, "heater 1", 5) Then
            _ErrorMsg = "Failed sending command 'heater'"
            Return False
        End If

        For Each Line In Response.Split(Chr(13), Chr(10))
            If Regex.IsMatch(Line, "^\d") Then
                DataLine = Line
            End If
        Next
        Form1.AppendText("Line = |" + DataLine + "|")
        Fields = Split(DataLine, ",")
        If Not Fields.Count = DataLineFields.Count Then
            _ErrorMsg = "Heater data line has " + Fields.Count.ToString + " fields, expecting " + DataLineFields.Count.ToString
            Return False
        End If

        Results = New Hashtable
        Try
            For i = 1 To UBound(DataLineFields)
                Name = Split(DataLineFields(i), "|")(0)
                Type = Split(DataLineFields(i), "|")(1)
                Select Case Type
                    Case "I"
                        Results.Add(Name, CInt(Fields(i)))
                    Case "D"
                        Results.Add(Name, CDbl(Fields(i)))
                    Case "S"
                        Results.Add(Name, Fields(i))
                    Case Else
                        _ErrorMsg = "Unknown type='" + Type + "' for parsing " + Name
                        Return False
                End Select
            Next

        Catch ex As Exception
            _ErrorMsg = "Problem parsing _DGA.PPM from " + Fields(1)
            Return False
        End Try

        Return True
    End Function

    Function SetPWM(ByVal pwm As String, ByVal setting As Double) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim PwmStatus As Hashtable

        If Not pwm = "OILBLOCK" And Not pwm = "ANALOGBD" Then
            _ErrorMsg = "Code Error:  SetPWM() pwm=" + pwm + ", expected pwm OILBLOCK or ANALOGBD"
            Return False
        End If

        If setting < 0 Or setting > 100 Then
            _ErrorMsg = "Code Error:  SetPWM() setting=" + setting.ToString + ", expected 0 - 100.0"
            Return False
        End If

        Command = "PWM " + pwm + " " + setting.ToString
        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "failed sending cmd " + Command
            Return False
        End If

        GetPwmStatus(PwmStatus)
        For Each k In PwmStatus.Keys
            Form1.AppendText(k + " " + PwmStatus(k)("RATE").ToString + " " + PwmStatus(k)("MODE"))
        Next
        If Not PwmStatus(pwm)("MODE") = "manual" Then
            _ErrorMsg = pwm + " " + PwmStatus(pwm)("RATE").ToString + " " + PwmStatus(pwm)("MODE") + vbCr +
                "Expected (manual)"
            Return False
        End If
        If Not PwmStatus(pwm)("RATE") = setting Then
            _ErrorMsg = pwm + " " + PwmStatus(pwm)("RATE").ToString + " " + PwmStatus(pwm)("MODE") + vbCr +
                "Expected " + setting.ToString
            Return False
        End If

        Return True
    End Function

    Function SetPWM_AUTO() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim PwmStatus As Hashtable

        Command = "PWM AUTO"
        If Not SF.Cmd(SerialPort, Response, Command, 5) Then
            _ErrorMsg = "failed sending cmd " + Command
            Return False
        End If

        GetPwmStatus(PwmStatus)
        For Each k In PwmStatus.Keys
            Form1.AppendText(k + " " + PwmStatus(k)("RATE").ToString + " " + PwmStatus(k)("MODE"))
        Next
        If Not PwmStatus("ANALOGBD")("MODE") = "auto" Then
            _ErrorMsg = "Expected AnalogBd mode=auto"
            Return False
        End If

        If Not PwmStatus("OILBLOCK")("MODE") = "auto" Then
            _ErrorMsg = "Expected OilBlock mode=auto"
            Return False
        End If

        Return True
    End Function

    Function GetPwmStatus(ByRef PwmStatus As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim Status As Hashtable
        Dim success As Boolean = False
        Dim retryCnt As Integer = 0

        While Not success And retryCnt < 3
            _ErrorMsg = ""
            If retryCnt > 0 Then
                CommonLib.Delay(2)
            End If
            retryCnt += 1
            System.Threading.Thread.Sleep(50)
            If Not SF.Cmd(SerialPort, Response, "pwm status", 5) Then
                _ErrorMsg = "failed sending cmd config"
                Return False
            End If
            PwmStatus = New Hashtable
            For Each Line In Response.Split(Chr(13), Chr(10))
                If Regex.IsMatch(Line, "(\w+)\s+PWM is\s+(\d+\.\d+)\s+\((\w+)\)") Then
                    Fields = Regex.Split(Line, "(\w+)\s+PWM is\s+(\d+\.\d+)\s+\((\w+)\)")
                    Status = New Hashtable
                    Try
                        Status.Add("RATE", CDbl(Fields(2)))
                    Catch ex As Exception
                        _ErrorMsg = "Could not extract pwm rate from " + Line
                        Return False
                    End Try
                    Status.Add("MODE", Fields(3))
                    PwmStatus.Add(Fields(1).ToUpper, Status)
                End If
            Next
            success = True
            For Each tec In {"OILBLOCK", "ANALOGBD"}
                If Not PwmStatus.Contains(tec) Then
                    _ErrorMsg = "no PWM status for " + tec
                    success = False
                    Exit For
                End If
                For Each param In {"RATE", "MODE"}
                    If Not PwmStatus(tec).contains(param) Then
                        _ErrorMsg += tec + " pwm status missing " + param + vbCr
                    End If
                Next
                If Not _ErrorMsg = "" Then
                    success = False
                    Exit For
                End If
            Next
        End While
        Return success
    End Function

    Function ReadADC(ByRef readings As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String

        readings = New Hashtable

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(SerialPort, Response, "adc 1", 10) Then
            _ErrorMsg = "Problem sending command 'adc 1'"
            Return False
        End If

        For Each Line In Response.Split(Chr(10), Chr(13))
            If Line.Contains("mA") Then
                Fields = Regex.Split(Line, "\s+")
                Dim fiveVolt As Integer = 7
                Dim twentyFourVolt As Integer = 5
                Dim aux1 As Integer = 1
                Dim aux2 As Integer = 3

                Select Case Fields.Count
                    Case 28
                        aux1 = 0
                        aux2 = 2
                        twentyFourVolt = 4
                        fiveVolt = 25

                    Case 29
                        fiveVolt = 26
                End Select

                If Not Fields.Count >= 9 Then
                    _ErrorMsg = "adc output has " + Fields.Count.ToString + "fields, expecting at least 9"
                    Return False
                End If
                Try
                    readings.Add("AUX1", CDbl(Regex.Split(Fields(aux1), "(.*)mA")(1)))
                    readings.Add("AUX2", CDbl(Regex.Split(Fields(aux2), "(.*)mA")(1)))
                    readings.Add("24VREF", CDbl(Regex.Split(Fields(twentyFourVolt), "(.*)v")(1)))
                    readings.Add("5VREF", CDbl(Regex.Split(Fields(fiveVolt), "(.*)v")(1)))
                Catch ex As Exception
                    _ErrorMsg = ex.ToString
                    Return False
                End Try
            End If
        Next

        Return True
    End Function

    Function GetSensors(ByRef SensorInfo As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim Name, Value As String

        SensorInfo = New Hashtable

        If Not SF.Cmd(SerialPort, Response, "sensor", 10) Then
            _ErrorMsg = "Problem sending command 'sensor'"
            Return False
        End If

        For Each Line In Response.Split(Chr(10), Chr(13))
            If Not Line.Contains(":") Then
                Continue For
            End If
            Line = Regex.Replace(Line, "\s+is\s+", ": ")
            Fields = Regex.Split(Line, ":")
            Try
                Name = Fields(0).Trim
                Value = Fields(1).Trim
                SensorInfo.Add(Name, Value)
            Catch ex As Exception
                _ErrorMsg = "Problem parsing sensor line" + Line
                Return False
            End Try
        Next
        Return True
    End Function

    Public Function ReadConsoleLog(ByRef console_log As ArrayList) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim console_log_found As Boolean
        Dim cat_console_success As Boolean
        Dim retryCnt As Integer

        console_log = New ArrayList
        console_log_found = False
        retryCnt = 0
        While (Not console_log_found And retryCnt < 3)
            retryCnt += 1
            CommonLib.Delay(2)
            If Not SF.Cmd(SerialPort, Response, "ls a:console_log", 15, Quiet:=True) Then
                _ErrorMsg = "problem running command 'ls a:console_log'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                CommonLib.Delay(10)
                Continue While
                ' Return False
            End If

            For Each Line In Response.Split(Chr(13))
                If Line.Contains("Number of files listed: 1") Then
                    console_log_found = True
                End If
            Next
        End While

        If Not console_log_found Then
            _ErrorMsg = "console_log file does not exist on the SD card"
            Return False
        End If

        cat_console_success = False
        retryCnt = 0
        While (Not cat_console_success And retryCnt < 3)
            If SF.Cmd(SerialPort, Response, "cat console_log", 20, Quiet:=True) Then
                cat_console_success = True
            End If
            retryCnt += 1
        End While
        If Not cat_console_success Then
            _ErrorMsg = "problem running command 'cat console_log'" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        For Each Line In Response.Split(Chr(13))
            Line = Line.Trim
            console_log.Add(Line)
        Next
        Return True
    End Function

    Function ls(ByRef file_listing As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim Name As String
        Dim Size As Integer


        file_listing = New Hashtable
        If Not SF.Cmd(SerialPort, Response, "ls", 10) Then
            _ErrorMsg = SF.ErrorMsg
            Return False
        End If
        For Each Line In Response.Split(Chr(13))
            Line = Line.Trim
            If Line.Contains("-rw") Then
                Fields = Regex.Split(Line, "\s+")
                Try
                    Name = Fields(8)
                    Size = CInt(Fields(4))
                    file_listing.Add(Name, Size)
                Catch ex As Exception
                    _ErrorMsg = "Problem parsing file info from " + Line + vbCr
                    _ErrorMsg += ex.ToString
                    Return False
                End Try
            End If
        Next

        Return True
    End Function

End Class
'                                                                                                                                                                                                                                                                                                                   