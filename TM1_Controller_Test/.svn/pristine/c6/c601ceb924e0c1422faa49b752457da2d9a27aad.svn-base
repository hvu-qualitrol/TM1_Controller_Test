Partial Class Tests
    Public Shared Function Test_SETDATE() As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim ZuluTime As String
        Dim Response As String
        Dim Cmd As String
        'Dim DateTimeStr As String
        'Dim ServerTime As DateTime
        'Dim UUT_DateTime As DateTime
        Dim T As New TM1

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        Dim startTime As DateTime = Now
        If Not CommonLib.GetLocalTime(ZuluTime, "ZULU") Then
            Form1.AppendText("Problem getting network time", True)
            Form1.AppendText(ZuluTime, True)
            Return False
        End If

        Cmd = "date -s " + ZuluTime
        If Not SF.Cmd(SerialPort, Response, Cmd, 5) Then
            Form1.AppendText("failed sending cmd " + Cmd, True)
            Return False
        End If

        CommonLib.Delay(10)
        If Not T.VerifyDate() Then
            Return False
        End If
        Return True

    End Function
End Class
