Imports System.IO.Ports
Imports System.Text.RegularExpressions

Partial Class Tests
    Public Shared Function Test_BOOT() As Boolean
        Dim ftdi_device As New FT232R
        Dim Comport As String
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim DB As New DB
        Dim T As New TM1
        Dim Tm1Config As Hashtable
        Dim ControllerBoardSN As String
        Dim NewLinkInfo As Hashtable
        Dim retryCnt As Integer
        Dim success As Boolean
        Dim file_listing As Hashtable
        Dim Response As String
        Dim FTDI_LOCKED As Boolean = False
        Dim exception_message As String = ""

        Dim serialNumber As String = Form1.Serial_Number_Entry.Text
        If Form1.AssemblyLevel_ComboBox.Text = "TM101_TM102" Then
            serialNumber = tm102Sn
            SerialPort = Nothing
        End If

        If IsNothing(SerialPort) Then
            Comport = "UNKNOWN"
            retryCnt = 0
            success = False
            If Not ftdi_device.LockFtdi() Then
                Form1.AppendText(ftdi_device.failure_message)
                Return False
            End If
            FTDI_LOCKED = True
            While (Not success And retryCnt < 3)
                If retryCnt > 0 Then
                    CommonLib.Delay(10)
                End If
                retryCnt += 1
                success = True
                If Not ftdi_device.FindComportForSN(serialNumber, Comport) Then
                    Form1.AppendText("Could not find FTDI device with SN " + serialNumber, True)
                    success = False
                End If
                If success Then
                    If Not Comport.StartsWith("COM") Then
                        success = False
                    End If
                End If
            End While
            Form1.AppendText("serial port = " + Comport, True)
            If Not success Then
                Form1.AppendText("Timeout trying to find comport")
                ftdi_device.UnlockFtdi()
                Return False
            End If
            Form1.Label_USB_COM.Text = Comport

            CommonLib.Delay(15)
            SerialPort = New SerialPort(Comport, 115200, 0, 8, 1)
            SerialPort.Handshake = Handshake.RequestToSend
        End If

        If Not SerialPort.IsOpen Then
            CommonLib.Delay(15)
            retryCnt = 0
            success = False
            While (Not success And retryCnt < 3)
                If retryCnt > 0 Then
                    Form1.AppendText("RETRY")
                    CommonLib.Delay(7)
                End If
                retryCnt += 1
                Form1.AppendText("Opening serial port", True)
                Try
                    SerialPort.Open()
                    success = True
                Catch ex As Exception
                    Form1.AppendText("Problem opening serial port")
                    exception_message = ex.ToString
                    'Form1.AppendText(ex.ToString)
                End Try
            End While
            If Not success Then
                If FTDI_LOCKED Then
                    ftdi_device.UnlockFtdi()
                End If
                Form1.AppendText(exception_message)
                Return False
            End If
        End If
        If FTDI_LOCKED Then
            ftdi_device.UnlockFtdi()
        End If

        Form1.AppendText("logging in to TM1", True)
        SerialPort.ReadTimeout = 1000
        results = SF.Login(SerialPort, True)
        If Not results.PassFail Then
            Form1.AppendText("Login failed", True)
            Form1.AppendText(results.Result, True)
            Return False
        End If
        Form1.AppendText("Logged in", True)

        'Check initial temperatures of the analog board and oil block
        'Threading.Thread.Sleep(500)
        If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Then
            If Not CheckInitTemperature(T) Then
                Form1.AppendText("Failed: Initial temperature condition")
                Return False
            End If
        End If

        ' Check lenght of records.dat file, fail and delete if 0.
        If Not T.ls(file_listing) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If
        Form1.AppendText("TM1 file listing")
        For Each name In file_listing.Keys
            Form1.AppendText(name + " " + file_listing(name).ToString)
        Next
        If Not file_listing.Contains("RECORDS.DAT") Then
            Form1.AppendText("RECORDS.DAT file missing from SD card")
            Return False
        End If
        If file_listing("RECORDS.DAT") = 0 Then
            If Not SF.Cmd(SerialPort, Response, "rm RECORDS.DAT", 10) Then
                Form1.AppendText("Problem sending cmd 'rm RECORDS.DAT'")
                Return False
            End If
            Form1.AppendText("RECORDS.DAT file length is 0, likely corrupted during initial creation")
            Form1.AppendText("The file has been removed, power cycle the board and re-run the test")
            Return False
        End If

        ' Done in case of TM101_TM102 level
        If Form1.AssemblyLevel_ComboBox.Text = "TM101_TM102" Then
            Return True
        End If

        ' Verify board.serial_number and add or update link table
        If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Or Form1.AssemblyLevel_ComboBox.Text = "W223_REWORK" Then
            If Not T.GetConfig(Tm1Config) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
            If Not Tm1Config.Contains("board.serial_number") Then
                Form1.AppendText("board.serial_number not found in TM1 config")
                Return False
            End If
            ControllerBoardSN = Form1.Controller_Board_SN_Entry.Text
            If Not Tm1Config("board.serial_number") = ControllerBoardSN Then
                Form1.AppendText("TM1 board.serial_number='" + Tm1Config("board.serial_number") + "', Expected " + ControllerBoardSN)
                Return False
            End If
            NewLinkInfo = New Hashtable
            NewLinkInfo.Add("TM1", TM1_SN)
            NewLinkInfo.Add("CONTROLLER_BOARD", ControllerBoardSN)
            NewLinkInfo.Add("POWER_SUPPLY", Form1.PowerSupply_SN_Entry.Text)
            NewLinkInfo.Add("PUMP", Form1.Pump_SN_Entry.Text)
            If Not Form1.ListBox_TestMode.Text = "DEBUG" Then
                If Not DB.AddLinkData(NewLinkInfo) Then
                    Form1.AppendText("Problem updating TM1_Link_Table")
                    Form1.AppendText(DB.ErrorMsg)
                    Return False
                End If
            End If

            '' Set temperature to 30C to use it as the precondition set point 
            '' for the thermal test happening later in the test flow   
            'Dim NewConfig As New Hashtable
            'Form1.AppendText("Config analogbd.setpoint and oilblock.setpoint to 30C.")
            'NewConfig.Add("analogbd.setpoint", "30.000")
            'NewConfig.Add("oilblock.setpoint", "30.000")
            'If Not T.SetConfig(NewConfig, RebootNeeded, True) Then
            '    Form1.AppendText(T.ErrorMsg)
            '    Return False
            'End If

            If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Then
                Form1.AppendText("Config analogbd.setpoint and oilblock.setpoint to 30C.")
                If Not PreConditionTempAndTMCOM1(T) Then
                    Form1.AppendText("Failed temperature preconditioning")
                    Return False
                End If
            End If

            ' Check to see if TMCOM1 can be logged in
            If Not CheckTMCOM1(Comport) Then
                Form1.AppendText("Failed to log in TMCOM1.")
                Return False
            End If

        End If

        Return True
    End Function

    Shared Function PreConditionTempAndTMCOM1(ByRef T As TM1) As Boolean
        Dim rebootNeeded As Boolean = True

        ' Set temperature to 30C to use it as the precondition set point 
        ' for the thermal test happening later in the test flow   
        Dim newConfig As New Hashtable
        newConfig.Add("analogbd.setpoint", "30.000")
        newConfig.Add("oilblock.setpoint", "30.000")

        ' Enable the TMCOM1
        newConfig.Add("CLI_OVER_TMCOM1.ENABLE", "TRUE")
        newConfig.Add("TMCOM1.PROTOCOL", "CLI")
        newConfig.Add("TMCOM1.MODE", "RS232")

        If Not T.SetConfig(newConfig, rebootNeeded, True) Then
            Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        Return True
    End Function

    Shared Function CheckInitTemperature(ByRef T As TM1) As Boolean
        Dim Line As String
        Dim Fields() As String

        SerialPort.Write("heater" + Chr(10))

        Dim done As Boolean = False
        Dim analogBdTemp As Double = 0
        Dim oilBlockTemp As Double = 0
        Dim count As Integer = 0
        While Not done And count < 10
            Try
                Line = SerialPort.ReadLine()
            Catch ex As Exception
                Line = ""
            End Try
            If Regex.IsMatch(Line, "^\d\d\d\d-") Then
                Line = Regex.Replace(Line, Chr(13), "")
                Line = Regex.Replace(Line, Chr(10), "")
                Fields = Split(Line, ",")
                If Fields.Count = 22 Then
                    oilBlockTemp = CDbl(Fields(12))
                    analogBdTemp = CDbl(Fields(15))
                    done = True
                    'Form1.AppendText("Initial OilBlockTemp = " + oilBlockTemp.ToString + "; initial AnalogBdTemp = " + analogBdTemp.ToString)
                End If
                'Else
                '    Form1.AppendText("Not matched: " + Line)
            End If
            count += 1
        End While

        If Not done Then
            Form1.AppendText("Failed to read temperatures.")
            Return False
        Else
            Form1.AppendText("Initial OilBlockTemp = " + oilBlockTemp.ToString + "; initial AnalogBdTemp = " + analogBdTemp.ToString)
        End If

        If oilBlockTemp > 33 Then
            Form1.AppendText("Failed: Initial OilBlockTemp " + oilBlockTemp.ToString + " > 33C")
            Return False
        End If
        If analogBdTemp > 33 Then
            Form1.AppendText("Failed: Initial AnalogBdTemp " + analogBdTemp.ToString + " > 33C")
            Return False
        End If

        ' Send Control-D: End of transmission
        SerialPort.Write(Chr(4) + Chr(10))
        done = False
        count = 0
        While Not done And count < 10
            Try
                Line = SerialPort.ReadLine()
            Catch ex As Exception
                Line = ""
            End Try
            If Not Line = "" Then
                Form1.AppendText("Line = " + Line)
                If Line.Contains(">") Then
                    done = True
                End If
            End If
            count += 1
        End While

        Return True
    End Function

    Shared Function CheckTMCOM1(ByVal mainComPort As String) As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim comPort As String
        Dim comPorts(Form1.Tmcom1ComboBox.Items.Count) As String
        Dim count As Integer = 0

        ' Build a list of COM ports, with the selected one first, to test
        comPorts(count) = Form1.Tmcom1ComboBox.Text
        For i As Integer = 0 To Form1.Tmcom1ComboBox.Items.Count - 1
            comPort = Form1.Tmcom1ComboBox.Items(i).ToString
            If comPort <> Form1.Tmcom1ComboBox.Text And comPort <> mainComPort Then
                count += 1
                comPorts(count) = comPort
            End If
        Next

        Dim status As Boolean = False
        For Each comPort In comPorts
            If IsNothing(comPort) Then Continue For

            Try
                Tmcom1SerialPort = New SerialPort(comPort, 115200, 0, 8, 1)
                Tmcom1SerialPort.Handshake = Handshake.RequestToSend
                Tmcom1SerialPort.Open()

                Form1.AppendText("Trying to log into TM1 via TMCOM1 " + comPort)
                Tmcom1SerialPort.ReadTimeout = 1000
                Tmcom1SerialPort.WriteTimeout = 200
                results = SF.Login(Tmcom1SerialPort)
                If Not results.PassFail Then
                    Form1.AppendText("Login TMCOM1 failed")
                    Form1.AppendText(results.Result + vbCr)
                    Tmcom1SerialPort.Close()
                    Continue For
                End If
                Form1.AppendText("Logged in TMCOM1 successfully")
                Tmcom1SerialPort.Close()
                status = True
                Form1.Tmcom1ComboBox.Text = comPort
                Exit For
            Catch ex As Exception
                Form1.AppendText("CheckTMCOM1() caught " + ex.Message + vbCr)
                Continue For
            End Try
        Next

        Return status

    End Function
End Class