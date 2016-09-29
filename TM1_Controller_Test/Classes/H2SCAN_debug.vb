Imports System.Text.RegularExpressions

Public Class H2SCAN_debug
    Public Enum H2SCAN_SI_MODE
        CLI = 0
        MODBUS = 1
    End Enum

    Public Enum H2SCAN_OP_MODE
        FIELD = 0
        LAB = 1
    End Enum

    Private _ErrorMsg As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function Open() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim PromptFound As Boolean = False
        Dim retryCnt As Integer = 0

        System.Threading.Thread.Sleep(200)
        If Not SF.Cmd(SerialPort, Response, "sensor -d", 10, "Press CTRL-D to break out of this mode.") Then
            _ErrorMsg = "Sent cmd 'sensor -d', did not see response 'Press CTRL-D to break out of this mode.'" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText(Response)

        SerialPort.Write(Chr(13) + Chr(10))

        While (Not PromptFound And retryCnt < 4)
            CommonLib.Delay(1)
            If SF.Cmd(SerialPort, Response, Chr(27), 10, "H2scan: ", False, True, False) Then
                PromptFound = True
            End If
            Form1.AppendText(Response)
            retryCnt += 1
            'If Response.Contains("Messages") Then
            '    SerialPort.Write(Chr(13) + Chr(10))
            'End If
        End While
        If Not PromptFound Then
            _ErrorMsg = "Didn't get H2 scan prompt 'H2scan: '"
            Return False
        End If
        CommonLib.Delay(1)
        If Not SF.Cmd(SerialPort, Response, "=serv" + Chr(13) + Chr(10), 5, "H2scan: ", False, True, False) Then
            _ErrorMsg = Response + vbCr + "Problem sending H2SCAN command '=serv'"
            Return False
        End If
        Form1.AppendText(Response)

        'If Not SF.Cmd(SerialPort, Response, Chr(4), 10, "> ", False, True, False) Then
        '    Form1.AppendText("Couldn't exit from sensor debug mode, expected to see '> '")
        '    Form1.AppendText(SF.ErrorMsg)
        '    Return False
        'End If
        Form1.AppendText(Response)

        Return True
    End Function

    Public Function Close() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        If Not SF.Cmd(SerialPort, Response, Chr(4), 10, "> ", False, True, False) Then
            _ErrorMsg = "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function GetOperatingMode(ByRef op_mode As H2SCAN_OP_MODE) As Boolean
        Dim T As New TM1
        Dim SensorInfo As Hashtable

        If Not T.GetSensors(SensorInfo) Then
            _ErrorMsg = T.ErrorMsg
            Return False
        End If
        If SensorInfo("Sensor Operating Mode").ToString.StartsWith("Field") Then
            op_mode = H2SCAN_OP_MODE.FIELD
        ElseIf SensorInfo("Sensor Operating Mode").ToString.StartsWith("Lab") Then
            op_mode = H2SCAN_OP_MODE.LAB
        Else
            _ErrorMsg = "Cant decode H2SCAN operating mode from '" + SensorInfo("Sensor Operating Mode") + "'"
            Return False
        End If

        Return True
    End Function

    Public Function SetOperatingMode(ByRef op_mode As H2SCAN_OP_MODE) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim current_mode As H2SCAN_OP_MODE

        If Not Open() Then
            _ErrorMsg += "H2SCAN_debug.Open() failed." + vbCr
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "m" + Chr(13) + Chr(10), 5, "Change (Y/N)? ", False, True, False) Then
            _ErrorMsg = "Problem sending h2scan m cmd, expected to see 'Change (Y/N)?  '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        If Response.Contains("enabled") Then
            current_mode = H2SCAN_OP_MODE.LAB
        ElseIf Response.Contains("disabled") Then
            current_mode = H2SCAN_OP_MODE.FIELD
        Else
            _ErrorMsg = "Cannot determine if lab mode is enabled or disabled"
            If Not SF.Cmd(SerialPort, Response, Chr(4), 20, "> ", False, True, False, True) Then
                _ErrorMsg += "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            Return False
        End If
        Form1.AppendText("current_mode = " + current_mode.ToString)

        Form1.AppendText("changing operating modes")
        If Not SF.Cmd(SerialPort, Response, "y" + Chr(13) + Chr(10), 20, "(Y/N)? ", False, True, False, True) Then
            _ErrorMsg += "Problem answering yes to change mode" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("changing operating modes")
        If Not SF.Cmd(SerialPort, Response, "y" + Chr(13) + Chr(10), 20, "(Y/N)? ", False, True, False, True) Then
            _ErrorMsg += "Problem answering yes to change mode" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        If Not SF.Cmd(SerialPort, Response, "n" + Chr(13) + Chr(10), 20, "(Y/N)? ", False, True, False, True) Then
            _ErrorMsg += "Problem answering no to change mode" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        If Not SF.Cmd(SerialPort, Response, "y" + Chr(13) + Chr(10), 20, "H2scan: ", False, True, False, True) Then
            _ErrorMsg += "Didn't see H2scan prompt" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, Chr(4), 20, "> ", False, True, False, True) Then
            _ErrorMsg += "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function GetSerialMode(ByRef si_mode As H2SCAN_SI_MODE) As Boolean
        Dim T As New TM1
        Dim SensorInfo As Hashtable

        If Not T.GetSensors(SensorInfo) Then
            _ErrorMsg = T.ErrorMsg
            Return False
        End If
        If SensorInfo("Sensor Model") = "" Or SensorInfo("Sensor Model") = "??" Then
            si_mode = H2SCAN_SI_MODE.CLI
        Else
            si_mode = H2SCAN_SI_MODE.MODBUS
        End If
        Return True
    End Function

    Public Function SetSerialMode(ByVal si_mode As H2SCAN_SI_MODE) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        If Not SF.Cmd(SerialPort, Response, "si 2" + Chr(13) + Chr(10) + Chr(4), 20, "> ", False, True, False) Then
            _ErrorMsg = "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Return True
    End Function

    Public Function VerifyModbusID() As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim ModbusID As Integer = 0

        If Not SF.Cmd(SerialPort, Response, "mi" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent mi command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response)
        For Each Line In Regex.Split(Response, "\r\n")
            If Regex.IsMatch(Line, "Modbus ID is \d+ Change") Then
                Try
                    ModbusID = CInt(Regex.Split(Line, "Modbus ID is (\d+) Change")(1))
                Catch ex As Exception
                    _ErrorMsg = "Problem extracting modbus ID from line '" + Line + "'"
                    _ErrorMsg += ex.ToString
                    Return False
                End Try
            End If
            Form1.AppendText("Line = " + Line)
        Next
        'Form1.AppendText("Response = " + Response)

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(SerialPort, Response, "n " + Chr(13) + Chr(10), 5, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem sending 'n' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Form1.AppendText("Modbus ID = " + ModbusID.ToString)
        If Not ModbusID = 1 Then
            _ErrorMsg = "Expecting MosbusID = 1"
            Return False
        End If


        Return True
    End Function
End Class
