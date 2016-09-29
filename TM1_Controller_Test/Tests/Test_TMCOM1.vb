Imports System.IO.Ports

Partial Class Tests
    Public Shared Function Test_TMCOM1() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "config -S set EXTERNAL_MODEM.ENABLE false", 10) Then
            Form1.AppendText("failed sending cmd config set EXTERNAL_MODEM.ENABLE false")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        Else
            Form1.AppendText(Response)
        End If

        If Not SF.Cmd(SerialPort, Response, "config -S set TMCOM1.PROTOCOL CLI", 10) Then
            Form1.AppendText("failed sending cmd config -S set TMCOM1.PROTOCOL CLI")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        Else
            Form1.AppendText(Response)
        End If

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(SerialPort, Response, "config -S set TMCOM1.MODE RS232", 10) Then
            Form1.AppendText("failed sending cmd config -S set TMCOM1.MODE RS232")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        Else
            Form1.AppendText(Response)
        End If

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(SerialPort, Response, "config set CLI_OVER_TMCOM1.ENABLE true", 10) Then
            Form1.AppendText("failed sending cmd config set CLI_OVER_TMCOM1.ENABLE true")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        Else
            Form1.AppendText(Response)
        End If

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
        Tmcom1SerialPort = New SerialPort(Form1.Tmcom1ComboBox.Text, 115200, 0, 8, 1)
        Tmcom1SerialPort.Handshake = Handshake.RequestToSend
        Try
            Tmcom1SerialPort.Open()
        Catch ex As Exception
            Form1.AppendText("Problem opening TMCOM1 serial port")
            Return False
        End Try

        Form1.AppendText("logging in to TM1 on TMCOM1" + vbCr)
        Tmcom1SerialPort.ReadTimeout = 1000
        results = SF.Login(Tmcom1SerialPort)
        If Not results.PassFail Then
            Form1.AppendText("Login failed" + vbCr)
            Form1.AppendText(results.Result + vbCr)
            Tmcom1SerialPort.Close()
            Return False
        End If
        Form1.AppendText("Logged in" + vbCr)
        Tmcom1SerialPort.Close()

        Return True
    End Function
End Class