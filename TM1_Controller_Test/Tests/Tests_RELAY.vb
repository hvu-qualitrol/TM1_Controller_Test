Partial Class Tests
    Public Shared Function Test_RELAY() As Boolean
        Dim Relay As Integer
        ' Dim DF As New DAQ_functions
        Dim RDC As New RemoteDaqClient
        Dim T As New TM1
        Dim RelayControl As String
        Dim DriveLevel As String
        Dim RS As RelayState
        Dim TestStatus As Boolean
        Dim RelayStatus As Hashtable
        Dim ExpectedRelayStatus As String
        Dim ExpectedRelayTerminalLevel As Integer
        Dim RelayName As String
        Dim RelayTerminals() As String = {"NO", "NC"}
        Dim RelayTerminalLevel As Integer
        Dim results As ReturnResults
        Dim SF As New SerialFunctions



        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not RDC.LockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
            Return False
        End If

        'If Not DF.InitializeDAQ Then
        '    Form1.AppendText(DF.GetErrorMsg, True)
        '    Return False
        'End If
        If Not RDC.InitializeDAQ Then
            Form1.AppendText(RDC.ErrorMsg, True)
            If Not RDC.UnlockDaq Then
                Form1.AppendText(RDC.ErrorMsg)
            End If
            Return False
        End If

        For Relay = 1 To 4
            Form1.AppendText("Setting relay " + Relay.ToString + " to OFF", True)
            If Not T.SetRelay(SerialPort, Relay, RelayState.RELAY_OFF) Then
                Form1.AppendText("..problem setting relay " + Relay.ToString + " OFF", True)
                If Not RDC.UnlockDaq Then
                    Form1.AppendText(RDC.ErrorMsg)
                End If
                Return False
            End If

            Form1.AppendText("Driving relay " + Relay.ToString + " on DAQ pin " + T.GetRelayConnection(Relay, "COM").ToString + " high ", True)
            'If Not DF.DrivePin(T.GetRelayConnection(Relay, "COM"), 1) Then
            '    Form1.AppendText(DF.GetErrorMsg, True)
            '    Return False
            'End If
            If Not RDC.DrivePin(T.GetRelayConnection(Relay, "COM"), 1) Then
                Form1.AppendText(RDC.ErrorMsg, True)
                If Not RDC.UnlockDaq Then
                    Form1.AppendText(RDC.ErrorMsg)
                End If
                Return False
            End If
        Next

        For i_relay = 1 To 4
            ' Verify all relays off
            Form1.AppendText("##############################################", True)
            For Each TestState In {"OFF:1", "OFF:0", "ON:1", "ON:0", "OFF:1"}
                RelayControl = (TestState.Split(":"))(0)
                DriveLevel = (TestState.Split(":"))(1)
                Form1.AppendText("Setting relay " + i_relay.ToString + " to " + RelayControl + ", driving relay common to " + TestState.ToString, True)
                If RelayControl = "ON" Then
                    RS = RelayState.RELAY_ON
                Else
                    RS = RelayState.RELAY_OFF
                End If
                If Not T.SetRelay(SerialPort, i_relay, RS) Then
                    Form1.AppendText("..problem setting relay " + i_relay.ToString + " " + RelayControl + " ", True)
                    Form1.AppendText(T.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
                'If Not DF.DrivePin(T.GetRelayConnection(i_relay, "COM"), DriveLevel) Then
                '    Form1.AppendText(DF.GetErrorMsg, True)
                '    Form1.AppendText(T.ErrorMsg)
                '    Return False
                'End If
                If Not RDC.DrivePin(T.GetRelayConnection(i_relay, "COM"), DriveLevel) Then
                    Form1.AppendText(RDC.ErrorMsg, True)
                    Form1.AppendText(T.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If

                TestStatus = True
                If Not T.GetRelays(SerialPort, RelayStatus) Then
                    Form1.AppendText("Problem getting relay status", True)
                    Form1.AppendText(T.ErrorMsg)
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
                For j_relay = 1 To 4
                    If j_relay = i_relay Then
                        ExpectedRelayStatus = RelayControl
                    Else
                        ExpectedRelayStatus = "OFF"
                    End If
                    RelayName = "Relay" + j_relay.ToString
                    Form1.AppendText(RelayName + " = " + RelayStatus(RelayName), True)
                    If Not RelayStatus(RelayName) = ExpectedRelayStatus Then
                        Form1.AppendText("...Expected relay status = " + ExpectedRelayStatus, True)
                        TestStatus = False
                    End If
                Next
                If Not TestStatus Then
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If

                ' Verify NC and NO relay pin levels
                TestStatus = True
                For j_relay = 1 To 4
                    For Each RelayTerminal In RelayTerminals
                        If i_relay = j_relay And ((RelayTerminal = "NC" And RelayControl = "OFF") Or (RelayTerminal = "NO" And RelayControl = "ON")) Then
                            ExpectedRelayTerminalLevel = DriveLevel
                        Else
                            ExpectedRelayTerminalLevel = 1
                        End If
                        RelayName = "Relay" + j_relay.ToString
                        'If Not DF.ReadBit(T.GetRelayConnection(j_relay, RelayTerminal), RelayTerminalLevel) Then
                        '    Form1.AppendText(DF.GetErrorMsg, True)
                        '    Return False
                        'End If
                        If Not RDC.ReadBit(T.GetRelayConnection(j_relay, RelayTerminal), RelayTerminalLevel) Then
                            Form1.AppendText(RDC.ErrorMsg, True)
                            If Not RDC.UnlockDaq Then
                                Form1.AppendText(RDC.ErrorMsg)
                            End If
                            Return False
                        End If
                        Form1.AppendText(RelayName + " terminal " + RelayTerminal + " level = " + RelayTerminalLevel.ToString, True)
                        If Not RelayTerminalLevel = ExpectedRelayTerminalLevel Then
                            Form1.AppendText("Expected:  " + ExpectedRelayTerminalLevel.ToString, True)
                            TestStatus = False
                        End If
                    Next
                Next
                If Not TestStatus Then
                    If Not RDC.UnlockDaq Then
                        Form1.AppendText(RDC.ErrorMsg)
                    End If
                    Return False
                End If
            Next
            'RTB.AppendText(vbCr)
        Next

        If Not RDC.UnlockDaq Then
            Form1.AppendText(RDC.ErrorMsg)
        End If

        Return True
    End Function
End Class