Partial Class Tests
    Public Shared Function TestReworkConfig() As Boolean
        Dim PassFlag As Boolean = False

        If Test_VERIFY_CONFIG() Then
            PassFlag = DoFactoryReset()
        End If

        Return PassFlag
    End Function

    Public Shared Function Test_VERIFY_CONFIG() As Boolean
        Dim config As New config
        Dim ConfigHash As Hashtable
        Dim ConfigNames As ICollection
        Dim ConfigNamesArray() As String
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim Tm1Config As Hashtable
        Dim Response As String
        Dim Command As String

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        ' Set board.serial_number if running the controller board test.
        'If Form1.AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" And Not Form1.ListBox_TestMode.Text = "DEBUG" Then
        If Form1.AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" Then
            Command = "config -S set board.serial_number " + Form1.Serial_Number_Entry.Text
            If Not SF.Cmd(SerialPort, Response, Command, 10) Then
                Form1.AppendText("Problem sending cmd '" + Command + "'")
                Form1.AppendText(SF.ErrorMsg)
                Return False
            End If
            CommonLib.Delay(5)
        End If

        If Not SF.Cmd(SerialPort, Response, "config reset factory", 10) Then
            Form1.AppendText("failed sending cmd 'config reset factory'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        Else
            Form1.AppendText(Response)
        End If
        System.Threading.Thread.Sleep(2000)

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        'if Not config.ReadExpectedConfig("default_config", ConfigHash) Then
        'If Not config.ReadExpectedConfig("default_config_1.1", ConfigHash) Then
        If Not config.ReadExpectedConfig("default_config_1.4", ConfigHash) Then
            Form1.AppendText(config.ErrorMsg)
            Return False
        End If
        If Not ConfigHash.Contains("board.serial_number") Then
            Form1.AppendText("default config missing 'board.serial_number'")
            Return False
        End If
        If Form1.AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" Then
            ConfigHash("board.serial_number") = Form1.Serial_Number
        Else
            ConfigHash("board.serial_number") = Form1.Controller_Board_SN_Entry.Text
        End If

        CommonLib.Delay(2)
        If Not T.GetConfig(Tm1Config) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        ConfigNames = ConfigHash.Keys
        ReDim ConfigNamesArray(ConfigHash.Count - 1)
        ConfigNames.CopyTo(ConfigNamesArray, 0)
        Array.Sort(ConfigNamesArray)

        For Each k In Tm1Config.Keys
            If Not ConfigHash.Contains(k) Then
                Form1.AppendText("Unexpected config:  " + k)
                Return False
            End If
        Next

        For Each k In ConfigNamesArray
            If Not Tm1Config.Contains(k) Then
                Form1.AppendText("Did not find config entry for " + k)
                Return False
            End If
            Form1.AppendText(k + " = " + Tm1Config(k))
            If Not Tm1Config(k) = ConfigHash(k) Then
                Form1.AppendText("   Expected:  " + ConfigHash(k))
                Return False
            End If
        Next

        Return True
    End Function

    Public Shared Function DoFactoryReset() As Boolean
        Dim config As New config
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        ' Perform a factory reset.
        CommonLib.Delay(30)
        Form1.AppendText("Clearing logs and records")
        If Not SF.Cmd(SerialPort, Response, "fr -A", 30, Quiet:=True) Then
            Form1.AppendText(Response)
            Form1.AppendText("failed sending cmd 'fr -A'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If
        Form1.AppendText(Response)
        System.Threading.Thread.Sleep(2000)

        Form1.AppendText("Reconnecting to ")
        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Form1.AppendText("Problem reconnecting to after config reset and reboot")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        CommonLib.Delay(30)
        Form1.AppendText("Verifying recs cleared")
        If Not SF.Cmd(SerialPort, Response, "rec", 5, Quiet:=True) Then
            Form1.AppendText(Response)
            Form1.AppendText("failed sending cmd 'rec'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        Return True
    End Function

End Class