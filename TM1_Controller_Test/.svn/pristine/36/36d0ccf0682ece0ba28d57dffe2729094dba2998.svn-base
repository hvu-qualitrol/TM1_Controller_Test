Imports TM1_Controller_Test.FT232R
Imports FTD2XX_NET

Partial Class Tests
    Public Shared Function Test_USB_SERIAL() As Boolean
        'Dim ftdi_device As New FT232R
        Dim comport As String
        Dim LocID As UInteger = 0
        Dim FT232R_device As New FT232R
        Dim ftdi_device As New FTDI
        Dim startTime As DateTime
        Dim FTDI_CYCLED As Boolean
        Dim BoardSN As String = Form1.Controller_Board_SN_Entry.Text
        ' Dim DR As DialogResult
        Dim FtdiDevices As Hashtable
        Dim FtdiDevCnt As Integer

        ' Askes if the test is part of the sub assembly and if so then to power up
        ' the monitor.
        If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Or Form1.AssemblyLevel_ComboBox.Text = "W223_REWORK" Then
            'DR = MessageBox.Show("Power up the TM1", Form1.Serial_Number_Entry.Text + ":  Power up?", MessageBoxButtons.YesNo)
            'If DR = DialogResult.No Then
            '    Form1.AppendText("Operator indicated problem powering up.")
            '    ' LedTimer.Stop()
            '    Return False
            'End If
            If Not YesNo("Disconnect USB cable from the TM1 then power up the TM1", "Power up?") Then
                Form1.AppendText("Operator indicated problem powering up.")
                Return False
            End If
            If Not YesNo("Connect USB cable to the TM1", "Connected?") Then
                Form1.AppendText("Operator indicated problem connecting USB cable.")
                Return False
            End If
            CommonLib.Delay(30)
        End If

        Form1.AppendText("Obtaining FTDI lock")
        If Not FT232R_device.LockFtdi() Then
            Form1.AppendText("Problem getting FTDI lock")
            Form1.AppendText(FT232R_device.failure_message)
            Return False
        End If

        ' ********************** DEBUGGING CODE *********** 
        ' This code needs to be removed before being sent to Mfg.


        ' *************************************************

        Form1.AppendText("Checking if FTDI device SN already set")
        If FT232R_device.FindComportForSN(Form1.Serial_Number_Entry.Text, comport, FtdiDevices, True) Then
            Form1.AppendText("TM1 USB ComPort = " + comport, True)
            CommonLib.Delay(1)
            If Not IsNothing(SerialPort) Then
                If Not SerialPort.PortName = comport Then
                    SerialPort = Nothing
                End If
            End If
            FT232R_device.UnlockFtdi()

            For Each SN In FtdiDevices.Keys
                Form1.AppendText("found ftdi SN=" + SN + ", comport=" + FtdiDevices(SN)("COM") + ", LocID=" + FtdiDevices(SN)("LOCID").ToString)
            Next
            Form1.AppendText("FTDI SN Already Set, skipping FTDI SN programming")
            Form1.Label_USB_COM.Text = comport
            Return True
        End If

        'For Each SN In FtdiDevices.Keys
        '    Form1.AppendText("found ftdi SN=" + SN + ", comport=" + FtdiDevices(SN)("COM") + ", LocID=" + FtdiDevices(SN)("LOCID").ToString)
        'Next

        ' If the FTDI device is not set to the TM1 SN, check if set to the board SN and fail if not.
        CommonLib.Delay(10)
        'FtdiMM_File.Dispose()
        'FtdiHash.Clear()
        If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Or Form1.AssemblyLevel_ComboBox.Text = "W223_REWORK" Then
            Form1.AppendText("Looking for FTDI set to " + BoardSN)
            If FtdiDevices.Contains(BoardSN) Then
                LocID = FtdiDevices(BoardSN)("LOCID")
            Else
                If Not FT232R_device.FindLocationForSN(BoardSN, LocID) Then
                    Form1.AppendText("FTDI device not found with SN set to " + BoardSN)
                    Form1.AppendText(FT232R_device.failure_message)
                    FT232R_device.UnlockFtdi()
                    ' Return False
                End If
            End If
        Else
            FtdiDevCnt = 0
            For Each SN In FtdiDevices.Keys
                If FtdiDevices(SN)("LOCID") < &H1000 And FtdiDevices(SN)("LOCID") > 0 Then
                    LocID = FtdiDevices(SN)("LOCID")
                    FtdiDevCnt += 1
                End If
            Next
            If Not FtdiDevCnt = 1 Then
                If Not FT232R_device.GetLocation(LocID) Then
                    Form1.AppendText("Problem getting ftdi device location")
                    Form1.AppendText(FT232R_device.failure_message, True)
                    FT232R_device.UnlockFtdi()
                    Return False
                End If
            End If
        End If

        ' In case of controller board swapping, try to search "TM1" or the port with low LocID
        If LocID = 0 Then
            Form1.AppendText("Last attempt to find the correct divice...")
            For Each SN In FtdiDevices.Keys
                If Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" And SN.ToString.Contains("TM1") Then
                    Form1.AppendText("found ftdi SN=" + SN + ", comport=" + FtdiDevices(SN)("COM") + ", LocID=" + FtdiDevices(SN)("LOCID").ToString)
                    LocID = FtdiDevices(SN)("LOCID")
                    Exit For
                ElseIf FtdiDevices(SN)("LOCID") < &H1000 And FtdiDevices(SN)("LOCID") > 0 Then
                    Form1.AppendText("found ftdi SN=" + SN + ", comport=" + FtdiDevices(SN)("COM") + ", LocID=" + FtdiDevices(SN)("LOCID").ToString)
                    LocID = FtdiDevices(SN)("LOCID")
                    Exit For
                End If
            Next
        End If

        ' Fail if cannot find the port
        If LocID = 0 Then
            FT232R_device.UnlockFtdi()
            Form1.AppendText("Failed to find the COM port for the specified FTDI device.")
            Return False
        End If

        Form1.AppendText("USB Location ID = 0x" + String.Format("{0:x}", LocID), True)
        If Not ftdi_device.OpenByLocation(LocID) = FTDI.FT_STATUS.FT_OK Then
            Form1.AppendText("Problem opening FTDI device at " + String.Format("{0:x}", LocID), True)
            FT232R_device.UnlockFtdi()
            Return False
        End If
        If Not FT232R_device.SetSN(ftdi_device, Form1.Serial_Number_Entry.Text) Then
            Form1.AppendText("Problem setting ftdi device SN")
            Form1.AppendText(FT232R_device.failure_message, True)
            FT232R_device.UnlockFtdi()
            Return False
        End If
        Form1.AppendText("SN set", True)
        If Not ftdi_device.CyclePort() = FTDI.FT_STATUS.FT_OK Then
            Form1.AppendText("Problem cycling USB port", True)
            FT232R_device.UnlockFtdi()
            Return False
        End If
        CommonLib.Delay(18)

        startTime = Now
        FTDI_CYCLED = False
        While Not FTDI_CYCLED And Now.Subtract(startTime).TotalSeconds < 30
            CommonLib.Delay(3)
            If ftdi_device.OpenBySerialNumber(Form1.Serial_Number_Entry.Text) = FTDI.FT_STATUS.FT_OK Then
                FTDI_CYCLED = True
            End If
        End While
        If Not FTDI_CYCLED Then
            Form1.AppendText("Problem opening FTDI device witn SN " + Form1.Serial_Number_Entry.Text, True)
            FT232R_device.UnlockFtdi()
            Return False
        End If
        Form1.AppendText("USB Cycle time = " + Math.Round(Now.Subtract(startTime).TotalSeconds).ToString, True)
        ftdi_device.Close()
        Form1.AppendText("Verified access to device with SN " + Form1.Serial_Number_Entry.Text, True)

        SerialPort = Nothing
        FT232R_device.UnlockFtdi()

        Return True
    End Function
End Class