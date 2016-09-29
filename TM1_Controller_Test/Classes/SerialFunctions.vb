Imports System.Text.RegularExpressions
Imports System.IO.Ports

Public Class SerialFunctions
    Private _ErrorMsg As String = ""

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property


    Public Function GetPrompt(ByVal SerialPort As Object, Optional ByVal Prompt As String = "> ") As Boolean
        Dim Data As Integer
        Dim PromptFound As Boolean = False
        Dim StartTime As DateTime = Now
        Dim Buffer As String
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim CommandTerm = Chr(10)

        If Prompt = "H2scan: " Then
            CommandTerm = Chr(13) + Chr(10)
        End If

        'SerialPort.ReadExisting()
        SerialPort.DiscardInBuffer()
        System.Threading.Thread.Sleep(50)
        'SerialPort.Write(Chr(10))
        SerialPort.Write(CommandTerm)
        While (Not PromptFound And Now.Subtract(StartTime).TotalSeconds < 2)
            Try
                Application.DoEvents()
                Data = SerialPort.ReadChar()
                Buffer += Chr(Data)
                If (Buffer.EndsWith(Prompt)) Then
                    PromptFound = True
                End If
            Catch ex As Exception
                'SerialPort.Write(Chr(13) + Chr(10))
                'Buffer = ""
                'System.Threading.Thread.Sleep(200)
            End Try
        End While

        If Not PromptFound Then
            results = SF.Login(SerialPort)
            If Not results.PassFail Then
                _ErrorMsg = "Did not see prompt"
                Return False
            End If
            System.Threading.Thread.Sleep(200)
        End If

        Return True
    End Function

    Public Function Cmd(ByVal SerialPort As Object, ByRef Response As String, ByVal Command As String, ByVal Timeout As Integer, Optional ByVal Prompt As String = "> ", Optional ByVal Quiet As Boolean = False,
                        Optional ByVal NoCR As Boolean = False, Optional ByVal GetPromptFirst As Boolean = True,
                        Optional ByVal DisplayResults As Boolean = False, Optional ByVal DisplayTimeout As Boolean = False)
        Dim PromptFound As Boolean
        Dim StartTime As DateTime
        Dim Data As Integer
        'Dim ResponseLines() As String
        Dim Line As String
        Dim LineCnt As Integer
        ' Dim Prompt = "> "
        Dim PostCommandInput As String = "NONE"
        Dim CommandTerm = Chr(10)

        If Prompt = "H2scan: " Then
            CommandTerm = Chr(13) + Chr(10)
        End If

        If GetPromptFirst Then
            If Not GetPrompt(SerialPort, Prompt) Then
                Return False
            End If
        End If
        'SerialPort.ReadExisting()
        SerialPort.DiscardInBuffer()

        If (Not Quiet) Then Form1.AppendText(Command + vbCr, True)
        'Form1.AppendText(Command + vbCr, True)
        'SerialPort.Write(Command + Chr(10))
        System.Threading.Thread.Sleep(50)
        For Each C As Char In Command.ToCharArray
            SerialPort.Write(C)
            System.Threading.Thread.Sleep(2)
        Next
        'SerialPort.Write(Command)
        'If Not NoCR Then SerialPort.Write(Chr(10))
        'If Not NoCR Then SerialPort.Write(Chr(13) + Chr(10))
        If Not NoCR Then SerialPort.Write(CommandTerm)
        PromptFound = False
        Response = ""
        'If Command = "adc" Then
        '    System.Threading.Thread.Sleep(50)
        '    SerialPort.Write(Chr(3))
        'End If
        'If Command = "heater" Then
        '    System.Threading.Thread.Sleep(100)
        '    SerialPort.Write(Chr(4))
        'End If
        If Command = "config reset factory" Then
            Prompt = "(y/n)? "
            PostCommandInput = "y"
        End If
        If Command = "config set CLI_OVER_TMCOM1.ENABLE true" Or Command = "fr -E" Or Command = "fr -A" Then
            Prompt = "reboot the system now? "
            PostCommandInput = "y"
        End If
        If Command = "reboot" Then
            PromptFound = True
            CommonLib.Delay(5)
        End If
        StartTime = Now
        ' Need to handle case where unit is not logged in
        While (Not PromptFound And Now.Subtract(StartTime).TotalSeconds < Timeout)
            Try
                If DisplayTimeout Then
                    Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(StartTime).TotalSeconds).ToString
                End If
                Application.DoEvents()
                Data = SerialPort.ReadChar()
                If DisplayResults Then
                    If Not Data = 10 And Not Data = 13 Then
                        Form1.AppendText(Chr(Data), True, False)
                    ElseIf Data = 13 Then
                        Form1.AppendText(vbCr, True, False)
                    End If
                End If
                Response += Chr(Data)
                If (Response.EndsWith(Prompt)) Then
                    PromptFound = True
                End If
            Catch ex As Exception

            End Try
        End While
        If Not PromptFound Then
            _ErrorMsg = "Did not see prompt:  " + Prompt
            Return False
        End If
        If Not PostCommandInput = "NONE" Then
            SerialPort.Write(PostCommandInput)
        End If
        'If Command = "config reset factory" Then
        '    SerialPort.Write("y")
        '    Return True
        'End If

        ' Trim command & prompt from response
        Response = Regex.Replace(Response, "^" + Command + ".*" + Chr(10), "")
        If Quiet Then
            Response = Regex.Replace(Response, Chr(13) + Chr(10) + Prompt + ".*$", "")
        End If
        Return True
    End Function


    Public Function Login(SP As Object, Optional ByVal WaitRecordsFile As Boolean = False) As ReturnResults
        Dim prompt_found As Boolean
        Dim Results As New ReturnResults
        Dim startTime As DateTime
        Dim Data As Integer
        Dim Buffer As String
        Dim Timeout As Integer = 10
        Dim NotLoggedIn As Boolean = False
        Dim DataReceivedTime As DateTime = Now

        Results.PassFail = False

        If WaitRecordsFile Then Timeout = 300

        If Not SP.IsOpen Then
            Results.PassFail = False
            Results.Result = "Serial port is not open"
            Return Results
        End If

        prompt_found = False
        startTime = Now
        'SP.RtsEnable = True
        While (Not prompt_found And Now.Subtract(startTime).TotalSeconds < Timeout And Not Stopped)
            Application.DoEvents()
            Try
                Data = SP.ReadChar()
                DataReceivedTime = Now
                If Not Data = 13 Then
                    Form1.AppendText(Chr(Data), True, False)
                    If WaitRecordsFile Then
                        Timeout = 10
                        WaitRecordsFile = False
                        startTime = Now
                    End If
                End If
                Buffer += Chr(Data)
                If (Buffer.EndsWith("username: ")) Then
                    System.Threading.Thread.Sleep(200)
                    SP.Write("BPLG" + Chr(10) + "CommandTimeout" + Chr(10))
                    NotLoggedIn = True
                End If
                If (Buffer.EndsWith("> ")) Then
                    prompt_found = True
                End If
                If (Buffer.Contains(": ") Or Buffer.Contains("> ")) Then
                    Buffer = ""
                End If

            Catch ex As Exception
                Try
                    SP.Write(Chr(13) + Chr(10))
                Catch exx As Exception
                    Form1.AppendText(exx.ToString)
                End Try
                Buffer = ""
                If Now.Subtract(DataReceivedTime).TotalSeconds > 10 Then
                    CommonLib.Delay(10)
                Else
                    System.Threading.Thread.Sleep(200)
                End If
                Form1.TimeoutLabel.Text = Int((Timeout - Now.Subtract(startTime).TotalSeconds)).ToString
            End Try
        End While
        Form1.TimeoutLabel.Text = "NA"
        Results.PassFail = prompt_found
        If NotLoggedIn Then System.Threading.Thread.Sleep(300)

        Return Results
    End Function

    Public Function Connect(ByRef SerialPort As Object) As ReturnResults
        Dim ftdi_device As New FT232R
        Dim Comport As String
        Dim Results As New ReturnResults

        Results.PassFail = False

        If IsNothing(SerialPort) Then
            If Not ftdi_device.LockFtdi() Then
                Form1.AppendText(ftdi_device.failure_message)
                Return Results
            End If
            If Not ftdi_device.FindComportForSN(Form1.Serial_Number_Entry.Text, Comport) Then
                Form1.AppendText("Could not find FTDI device with SN " + Form1.Serial_Number_Entry.Text, True)
                ftdi_device.UnlockFtdi()
                Return Results
            End If
            Form1.AppendText("serial port = " + Comport, True)

            CommonLib.Delay(15)
            SerialPort = New SerialPort(Comport, 115200, 0, 8, 1)
            SerialPort.Handshake = Handshake.RequestToSend

            ' Trying with this encoding
            'SerialPort.Encoding = System.Text.Encoding.GetEncoding(28591)
            'SerialPort.Encoding = System.Text.Encoding.GetEncoding("iso-8859-1")
            ftdi_device.UnlockFtdi()
        End If

        If Not SerialPort.IsOpen Then
            Form1.AppendText("Opening serial port", True)
            Try
                SerialPort.Open()
            Catch ex As Exception
                Form1.AppendText("Problem opening serial port", True)
                Return Results
            End Try
        End If

        Form1.AppendText("logging in to TM1", True)
        SerialPort.ReadTimeout = 1000
        Results = Login(SerialPort)
        If Not Results.PassFail Then
            Form1.AppendText("Login failed", True)
            Form1.AppendText(Results.Result, True)
            Return Results
        End If
        ' Form1.AppendText("Logged in", True)

        Return Results
    End Function

    Public Function Reboot(ByVal SerialPort As Object) As ReturnResults
        Dim Results As New ReturnResults
        Dim Response As String

        Results.PassFail = False
        If Not SerialPort.IsOpen Then
            Results.PassFail = False
            Results.Result = "Serial port is not open"
            Return Results
        End If
        If Not Cmd(SerialPort, Response, "reboot", 10) Then
            Return Results
        End If
        Results = Login(SerialPort)

        Return Results
    End Function

    Public Function Close() As Boolean
        If Not SerialPort Is Nothing Then
            If SerialPort.IsOpen Then
                Try
                    SerialPort.Close()
                    SerialPort = Nothing
                Catch ex As Exception

                End Try
            End If
        End If

        Return True
    End Function


End Class
