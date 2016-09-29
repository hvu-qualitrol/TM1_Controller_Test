Imports System.IO.Ports

''' <summary>
''' This partial class module is to provide needed functions to test
''' TM1 Retrofit Display Kits
''' Test Flow:
''' Prompt the user to make the needed connection setups
''' Open the COM port connecting to the TM1
''' Log in the TM1
''' Perform tests under auto mode.
''' Perform tests under manual mode.
''' Close the commnunications
''' Prompt the user to disconnect the parts
''' </summary>
Partial Class Tests
    ''' <summary>
    ''' This function is to test TM1 Retrofit Display Kit. It return True/False
    ''' based on the pass/fail test status, respectively.
    ''' </summary>
    ''' <returns>True/False for pass/fail status</returns>
    Public Shared Function TestDisplay() As Boolean
        Dim passFlag As Boolean = True
        Form1.AppendText("TestDisplay is being invoked...")

        ' Prompt the user to make the needed setups
        Dim title As String = "TM1 Display Kit Test"
        Dim message As String = "Power down the TM1, connect the display kit, then power up the TM1. Click OK when ready."
        Dim result As DialogResult
        result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            Return False
        End If

        ' Try to open TM1 main COM port
        If AssemblyLevel = "DISPLAY_KIT" Then
            If Not OpenMainSerialPort() Then
                Form1.AppendText("Problem opening TM1 Main serial port " + Form1.Tmcom1ComboBox.Text)
                If SerialPort.IsOpen() Then
                    SerialPort.Close()
                End If
                Return False
            End If
        End If

        ' Try to log in the TM1
        Dim SF As New SerialFunctions
        Dim results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            SF.Close()
            Return False
        End If

        ' Get embedded moiture option
        Dim VersionInfo As Hashtable = Nothing
        Dim T As New TM1
        If Not T.GetVersionInfo(VersionInfo) Then
            Return False
        End If

        ' Verify the display in auto mode
        If Not VerifyAutoDisplay(SF, VersionInfo("embedded moisture")) Then
            Form1.AppendText("Failed VerifyAutoDisplay test.")
            SF.Close()
            Return False
        End If

        ' Verify the display in manual mode
        If Not VerifyManualDisplay(SF) Then
            Form1.AppendText("Failed VerifyManualDisplay test.")
            SF.Close()
            Return False
        End If

        ' Link the display kit to the TM1 unit
        If Not LinkDisplay(TM1_SN, Form1.Controller_Board_SN_Entry.Text) Then
            Form1.AppendText("Failed LinkDisplay test.")
            SF.Close()
            Return False
        End If

        ' Clean up
        EndTest(SF)

        Form1.AppendText("TestDisplay is complete.")

        Return passFlag
    End Function

    Private Shared Function LinkDisplay(ByVal tm1Sn As String, ByVal displaySn As String) As Boolean
        Dim DB As New DB
        Dim NewLinkInfo As Hashtable
        NewLinkInfo = New Hashtable
        NewLinkInfo.Add("TM1", tm1Sn)
        NewLinkInfo.Add("DISPLAY_KIT", displaySn)
        If Not Form1.ListBox_TestMode.Text = "DEBUG" Then
            If Not DB.AddLinkData(NewLinkInfo) Then
                Form1.AppendText("Problem linking DISPLAY_KIT to TM1_Link_Table")
                Form1.AppendText(DB.ErrorMsg)
                Return False
            End If
        End If

        Return True
    End Function

    ''' <summary>
    ''' This private function is to send the messages to the display kit.
    ''' </summary>
    ''' <param name="SF">The reference to the SerialFunctions</param>
    ''' <param name="l1Msg">The message to be displayed on the first line of the display</param>
    ''' <param name="l2Msg">The message to be displayed on the second line of the display</param>
    ''' <returns>True/False for pass/fail status</returns>
    Private Shared Function SendDisplayMessages(ByRef SF As SerialFunctions, ByVal l1Msg As String, ByVal l2Msg As String) As Boolean
        Dim response As String = ""
        Dim command As String = ""

        ' Send the l1Msg to the display line 1
        command = "vfd -l 1 -s '" + l1Msg + "'"
        If Not SF.Cmd(SerialPort, response, command, 10) Then
            Form1.AppendText("Failed to execute command " + command)
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        ' Send the l2Msg to the display line 2
        command = "vfd -l 2 -s '" + l2Msg + "'"
        If Not SF.Cmd(SerialPort, response, command, 10) Then
            Form1.AppendText("Failed to execute command " + command)
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        Return True
    End Function
    Private Shared displayDone As Boolean = False
    Public Class DisplayParameters
        Public Form1 As Form1
        Public SF As SerialFunctions
        Public l1Msg As String
        Public l2Msg As String
    End Class

    ''' <summary>
    ''' This private function is to send the messages to the display kit.
    ''' </summary>
    ''' <param name="parameters">The specified object parameters</param>
    Public Shared Sub DisplayMessages(ByVal parameters As DisplayParameters)
        displayDone = False

        While (Not displayDone)
            Dim response As String = ""
            Dim command As String = ""

            ' Send the l1Msg to the display line 1
            command = "vfd -l 1 -s '" + parameters.l1Msg + "'"
            'parameters.Form1.AppendText("Executing command: " + command)
            If Not parameters.SF.Cmd(SerialPort, response, command, 10) Then
                parameters.Form1.AppendText("Failed to execute command " + command)
                parameters.Form1.AppendText(parameters.SF.ErrorMsg)
                displayDone = True
                Exit Sub
            End If

            ' Send the l2Msg to the display line 2
            command = "vfd -l 2 -s '" + parameters.l2Msg + "'"
            'parameters.Form1.AppendText("Executing command: " + command)
            If Not parameters.SF.Cmd(SerialPort, response, command, 10) Then
                parameters.Form1.AppendText("Failed to execute command " + command)
                parameters.Form1.AppendText(parameters.SF.ErrorMsg)
                displayDone = True
                Exit Sub
            End If

            ' Pause for 50ms
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        End While
    End Sub

    ''' <summary>
    ''' This private funtion is to read the most recent data record and then to 
    ''' extract the H2Oil and H2Roc values from the record.
    ''' </summary>
    ''' <param name="SF">The reference to the SerialFunctions</param>
    ''' <param name="h2OilPpm">The reference to the return value of h2OilPpm</param>
    ''' <param name="h2RocPpm">The reference to the return value of h2RocPpm</param>
    ''' <returns>True/false for pass/fail status</returns>
    ''' <remarks>Use -1111 and -1111.1 in case of NaN</remarks>
    Private Shared Function ExtractDisplayData(ByRef SF As SerialFunctions, ByRef h2OilPpm As Integer, ByRef h2RocPpm As Double) As Boolean
        Dim response As String = ""
        Dim fields() As String = {}
        Dim values() As String = {}
        Dim command As String = "rec 1"
        Dim found As Boolean = False

        For i As Integer = 0 To 3
            ' Read the most recent (last) record
            SF.Cmd(SerialPort, response, command, 10)

            ' Check whether we got valid record data
            If response.Contains("H2_OIL.PPM") And fields.Contains("H2_ROC") Then
                found = True
                Exit For
            End If
        Next

        ' Failed if we could not find the interested parameters
        If found = False Then Return False

        fields = Split(response, ",")
        For Each field In fields
            If field.Contains("H2_OIL.PPM") Then
                values = Split(field, "=")
                If Not Integer.TryParse(values(1), h2OilPpm) Then
                    h2OilPpm = -1111
                End If
            ElseIf field.Contains("H2_ROC") Then
                values = Split(field, "=")
                If Not Double.TryParse(values(1), h2RocPpm) Then
                    h2RocPpm = -1111.1
                End If
                Exit For
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' This private function is to be called at the end of the test to reset the display,
    ''' to close the serial communications, and to prompt the user to disconnect the parts
    ''' </summary>
    ''' <param name="SF">The reference to the SerialFunctions object</param>
    ''' <returns>True/false for pass/fail status</returns>
    ''' <remarks></remarks>
    Private Shared Function EndTest(ByRef SF As SerialFunctions) As Boolean
        ' Restore the auto mode for the display
        Dim response As String = ""
        If Not SF.Cmd(SerialPort, response, "vfd auto", 10) Then
            Form1.AppendText("Failed to execute command 'vfd auto'")
            Form1.AppendText(SF.ErrorMsg)
        End If

        ' Close the serial communications
        ' SF.Close()

        '' Prompt the user to power down and disconnect the setups
        'Dim title As String = "TM1 Display Kit Test"
        'Dim message As String = "Power down the TM1 and disconnect the display kit. Click OK when ready."
        'Dim result As DialogResult
        'result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
        'If result = DialogResult.Cancel Then
        '    Return False
        'End If

        Return True
    End Function

    ''' <summary>
    ''' This private sub is to provide timer elapsing count for the specified timeout.
    ''' The specified message is displayed along with the count
    ''' </summary>
    ''' <param name="timeout">The specified timeout in seconds</param>
    ''' <param name="message">The specified display message</param>
    ''' <remarks></remarks>
    Private Shared Sub WaitTimeout(ByVal timeout As Double, ByVal message As String)
        Dim startTime As DateTime = Now
        While (Now.Subtract(startTime).TotalSeconds < timeout)
            Form1.TimeoutLabel.Text = message +
                Math.Round(timeout - Now.Subtract(startTime).TotalSeconds).ToString
            Application.DoEvents()
        End While
    End Sub

    ''' <summary>
    ''' This private function is to test the display kit functionalitys under auto mode.
    ''' </summary>
    ''' <param name="SF">The reference to SerialFunctions object</param>
    ''' <returns>True/false for pass/fail status</returns>
    ''' <remarks>Translate to NaN when seeing H2Roc special value of -1111.1.
    ''' </remarks>
    Private Shared Function VerifyAutoDisplay(ByRef SF As SerialFunctions, ByVal embeddedMoisture As String) As Boolean
        Dim title As String = "TM1 Display Kit Test"
        Dim message As String = ""
        Dim result As DialogResult
        Dim h2OilPpm As Integer
        Dim h2RocPpm As Double

        ' Read the H2_OIL.PPM and H2_ROC.PPM
        ExtractDisplayData(SF, h2OilPpm, h2RocPpm)

        ' Ask the user to turn on the display
        message = "If the display is off, press ON to turn it on. Click OK when ready."
        result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            Return False
        End If

        ' Ask the user to verify the auto H2 display
        Dim question As String = "Does the displayed messages match the following text?"
        Dim l1Msg As String = "H2  " + h2OilPpm.ToString() + "  ppm"
        Dim l2Msg As String = "H2 ROC  " + h2RocPpm.ToString("0.0") + "  ppm/d"
        If h2RocPpm = -1111.1 Then
            l2Msg = "H2 ROC  " + "NaN" + "  ppm/d"
        End If
        result = MessageBox.Show(question + vbNewLine + l1Msg + vbNewLine + l2Msg, title, MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return False
        End If

        ' Ask the user to verify the auto moisture display if applicable
        If embeddedMoisture = "yes" Then
            Dim rh, temp, ppm As Double
            If Not GetEmbeddedMoistureData(SF, rh, temp, ppm) Then
                Return False
            End If
            question = "Does the displayed messages match the following text?"
            l1Msg = "MOISTURE  " + ppm.ToString() + "  ppm"
            result = MessageBox.Show(question + vbNewLine + l1Msg + vbNewLine + l2Msg, title, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Return False
            End If
        End If

        ' Wait for 45s timeout
        Form1.AppendText("The display will automatically be off after 45s timeout. Waiting for timeout...")
        WaitTimeout(45.0, "The display will be off in ")

        ' Ask the user to verify whether the display is off after 45s
        question = "Is the display automatically off after 45s?"
        result = MessageBox.Show(question, title, MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return False
        End If

        ' Ask the user to again turn on the display
        message = "Press the display ON button to enable the display. Click OK when ready."
        result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            Return False
        End If

        ' Ask the user to verify the auto display
        question = "Does the displayed messages match the following text?"
        result = MessageBox.Show(question + vbNewLine + l1Msg + vbNewLine + l2Msg, title, MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' This private function is to test the display kit in manual mode. The designed
    ''' messages are one by one sent to the display and ask the user to verify
    ''' </summary>
    ''' <param name="SF">The reference to SerialFunctions object</param>
    ''' <returns>True/false for pass/fail status</returns>
    ''' <remarks>In manual mode, the display will be off (actually, it is in
    ''' reset mode, so there is no message to display) when turned on.
    ''' </remarks>
    Private Shared Function VerifyManualDisplay0(ByRef SF As SerialFunctions) As Boolean
        Dim passFlag As Boolean = True
        Dim response As String = ""

        ' Put the display in manual mode
        If Not SF.Cmd(SerialPort, response, "vfd manual", 10) Then
            Form1.AppendText("Failed to execute command 'vfd manual'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        ' Build the test messages
        Dim l1TestMessages As String() = {"This is the 1st line",
                                        "XXXXXXXXXXXXXXXXXXXX",
                                        "IIIIIIIIIIIIIIIIIIII",
                                        "OOOOOOOOOOOOOOOOOOOO",
                                        "88888888888888888888",
                                        "@@@@@@@@@@@@@@@@@@@@",
                                        "&&&&&&&&&&&&&&&&&&&&"}
        Dim l2TestMessages As String() = {"This is the 2nd line",
                                        "IIIIIIIIIIIIIIIIIIII",
                                        "XXXXXXXXXXXXXXXXXXXX",
                                        "88888888888888888888",
                                        "OOOOOOOOOOOOOOOOOOOO",
                                        "&&&&&&&&&&&&&&&&&&&&",
                                        "@@@@@@@@@@@@@@@@@@@@"}

        Dim title As String = "TM1 Display Kit Test"
        Dim question As String = "Does the displayed messages match the following text?"
        Dim message As String = "If the display is off, press ON to turn it on. Click OK when ready."
        Dim result As DialogResult
        Dim parameters As New Object()
        parameters(0) = SF
        For i As Integer = 0 To l1TestMessages.Length - 1
            For j As Integer = 0 To 3
                ' Ask the user to turn on the display
                result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
                If result = DialogResult.Cancel Then
                    Return False
                End If

                ' Send the messages to the display
                SendDisplayMessages(SF, l1TestMessages(i), l2TestMessages(i))

                ' Prompt the user to verify the displayed messages
                result = MessageBox.Show(question + vbNewLine + l1TestMessages(i) + vbNewLine + l2TestMessages(i), title, MessageBoxButtons.YesNo)
                If result = DialogResult.No And j = 3 Then
                    passFlag = False
                    Exit For
                ElseIf result = DialogResult.Yes Then
                    Exit For
                End If
            Next
        Next

        Return passFlag
    End Function

    Private Shared Function VerifyManualDisplay(ByVal SF As SerialFunctions) As Boolean
        Dim passFlag As Boolean = True
        Dim response As String = ""

        ' Put the display in manual mode
        If Not SF.Cmd(SerialPort, response, "vfd manual", 10) Then
            Form1.AppendText("Failed to execute command 'vfd manual'")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        ' Build the test messages
        Dim l1TestMessages As String() = {"This is the 1st line",
                                        "XXXXXXXXXXXXXXXXXXXX",
                                        "IIIIIIIIIIIIIIIIIIII",
                                        "OOOOOOOOOOOOOOOOOOOO",
                                        "88888888888888888888",
                                        "@@@@@@@@@@@@@@@@@@@@",
                                        "&&&&&&&&&&&&&&&&&&&&"}
        Dim l2TestMessages As String() = {"This is the 2nd line",
                                        "IIIIIIIIIIIIIIIIIIII",
                                        "XXXXXXXXXXXXXXXXXXXX",
                                        "88888888888888888888",
                                        "OOOOOOOOOOOOOOOOOOOO",
                                        "&&&&&&&&&&&&&&&&&&&&",
                                        "@@@@@@@@@@@@@@@@@@@@"}

        Dim title As String = "TM1 Display Kit Test"
        Dim question As String = "Does the displayed messages match the following text?"
        Dim message As String = "If the display is off, press ON to turn it on. Click OK when ready."
        Dim result As DialogResult
        Dim parameters As New DisplayParameters()
        parameters.Form1 = Form1
        parameters.SF = SF
        parameters.l1Msg = l1TestMessages(0)
        parameters.l2Msg = l2TestMessages(0)
        Dim displayThread As New Threading.Thread(AddressOf DisplayMessages)
        displayThread.IsBackground = True
        displayThread.Start(parameters)

        For i As Integer = 0 To l1TestMessages.Length - 1
            ' Ask the user to turn on the display
            result = MessageBox.Show(message, title, MessageBoxButtons.OKCancel)
            If result = DialogResult.Cancel Then
                Return False
            End If

            ' Load the new messages to send to the display
            parameters.l1Msg = l1TestMessages(i)
            parameters.l2Msg = l2TestMessages(i)

            ' Prompt the user to verify the displayed messages
            result = MessageBox.Show(question + vbNewLine + l1TestMessages(i) + vbNewLine + l2TestMessages(i), title, MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                passFlag = False
                Exit For
                'ElseIf result = DialogResult.Yes Then
                '    Exit For
            End If
        Next

        ' Wait for the thread termination
        displayDone = True
        displayThread.Join()

        Return passFlag
    End Function

    ''' <summary>
    ''' This function is to set up and open the main serial port of the TM1 used to
    ''' test the retrofit display kit.
    ''' </summary>
    ''' <returns>True/false for pass/fail</returns>
    ''' <remarks>Close the port in case it has been openned before re-open it
    ''' </remarks>
    Shared Function OpenMainSerialPort() As Boolean
        ' Don't bother with the null string
        If Form1.Tmcom1ComboBox.Text = vbNullString Then
            Return True
        End If

        SerialPort = New SerialPort(Form1.Tmcom1ComboBox.Text, 115200, 0, 8, 1)
        SerialPort.Handshake = Handshake.RequestToSend
        If SerialPort.IsOpen() Then
            SerialPort.Close()
        End If

        Try
            SerialPort.Open()
        Catch ex As Exception
            Form1.AppendText("Caught exception: " + ex.Message)
            Return False
        End Try

        Return True

    End Function

End Class