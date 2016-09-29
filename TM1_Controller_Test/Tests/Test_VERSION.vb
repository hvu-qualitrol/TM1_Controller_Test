Imports System.Text.RegularExpressions

Partial Class Tests
    Public Shared Function Test_VERSION() As Boolean
        Dim SF As New SerialFunctions
        Dim T As New TM1
        'Dim Response As String
        Dim results As ReturnResults
        'Dim Line As String
        Dim VersionInfo As Hashtable
        Dim FW_Version As String = "UNKNOWN"
        Dim SensorInfo As Hashtable

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not T.GetVersionInfo(VersionInfo) Then
            Return False
        End If
        If Not VersionInfo.Contains("firmware version") Then
            Form1.AppendText("Did not see 'firmware version' in output of ver cmd", True)
            Return False
        End If
        Try
            FW_Version = Regex.Split(VersionInfo("firmware version"), "\s+")(1)
        Catch ex As Exception
            Form1.AppendText("Problem extract fw version from " + VersionInfo("fimrware version"), True)
            Return False
        End Try
        If Not FW_Version = Products(Product)("FW_VERSION") Then
            Form1.AppendText("Expected 'firmware version' = " + Products(Product)("FW_VERSION"), True)
            Return False
        End If

        If Not VersionInfo.Contains("hardware version") Then
            Form1.AppendText("Did not see 'hardwave version' in output of ver cmd", True)
            Return False
        End If
        If (VersionInfo("hardware version") <> Products(Product)("hardware version 0") And
            VersionInfo("hardware version") <> Products(Product)("hardware version 1") And
            VersionInfo("hardware version") <> Products(Product)("hardware version 2")) Then
            Form1.AppendText("Expected 'hardware version' = " + Products(Product)("hardware version 0") +
                             " or " + Products(Product)("hardware version 1") +
                             " or " + Products(Product)("hardware version 2"), True)
            Return False
        End If

        If Not VersionInfo.Contains("assembly version") Then
            Form1.AppendText("Did not see 'assembly version' in output of ver cmd", True)
            Return False
        End If
        If (VersionInfo("assembly version") <> Products(Product)("assembly version 0") And
            VersionInfo("assembly version") <> Products(Product)("assembly version 1") And
            VersionInfo("assembly version") <> Products(Product)("assembly version 2")) Then
            Form1.AppendText("Expected 'assembly version' = " + Products(Product)("assembly version 0") +
                             " or " + Products(Product)("assembly version 1") +
                             " or " + Products(Product)("assembly version 2"), True)
            Return False
        End If

        If AssemblyLevel = "SUB_ASSEMBLY" Or AssemblyLevel = "W223_REWORK" Then
            If Not VersionInfo.Contains("sensor firmware version") Then
                Form1.AppendText("Did not see 'sensor firmware version' in output of ver cmd", True)
                Return False
            End If
            If Not VersionInfo("sensor firmware version") = Products(Product)("sensor firmware version") Then
                ' For rework units, either 3.35B or 3.36B would be OK
                If AssemblyLevel = "W223_REWORK" Then
                    If VersionInfo("sensor firmware version") <> "3.35B" Then
                        Form1.AppendText("Expected 'sensor firmware version' = " + Products(Product)("sensor firmware version") + " or 3.35B", True)
                        Return False
                    End If
                Else
                    Form1.AppendText("Expected 'sensor firmware version' = " + Products(Product)("sensor firmware version"), True)
                    Return False
                End If

                If Not T.GetSensors(SensorInfo) Then
                    Form1.AppendText(T.ErrorMsg)
                    Return False
                End If
            End If
        End If

        Return True
    End Function
End Class