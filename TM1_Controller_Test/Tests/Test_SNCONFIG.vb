Imports Microsoft.Office.Interop

Partial Class Tests
    ' Variables used for TM101->TM102 conversion
    Private Shared tm101Sn As String
    Private Shared tm102Sn As String

    Public Shared Function Test_SNCONFIG() As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Command As String
        Dim T As New TM1
        Dim VersionInfo As Hashtable
        Dim DB As New DB

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not T.GetVersionInfo(VersionInfo) Then
            Return False
        End If
        If Not (VersionInfo.Contains("serial number")) Then
            Return False
        End If
        If VersionInfo("serial number") = TM1_SN Then
            Form1.AppendText("serial number already set")
            Return True
        End If

        Command = "config -S set SERIAL_NUMBER$ " + TM1_SN
        If Not SF.Cmd(SerialPort, Response, "config -S set SERIAL_NUMBER$ " + TM1_SN, 10) Then
            Form1.AppendText("Problem sending cmd '" + Command + "'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        System.Threading.Thread.Sleep(150)
        results = SF.Reboot(SerialPort)
        If Not results.PassFail Then
            Form1.AppendText("Problem rebooting")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        If Not T.GetVersionInfo(VersionInfo) Then
            Return False
        End If
        If Not (VersionInfo.Contains("serial number")) Then
            Return False
        End If
        If Not VersionInfo("serial number") = TM1_SN Then
            Form1.AppendText("Expected 'serial number' = " + TM1_SN)
            Return False
        End If

        Return True
    End Function

End Class