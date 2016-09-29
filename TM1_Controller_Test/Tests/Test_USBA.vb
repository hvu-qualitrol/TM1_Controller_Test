Imports System.IO.File
Imports System.IO

Partial Class Tests
    Public Shared Function Test_USBA() As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim U As New USB
        Dim USB_DRIVE As String
        ' Dim DR As DialogResult
        Dim ExpectedStrings As ArrayList
        Dim ExpectedStringItem As Hashtable
        Dim ExpectedStringIndex As Integer
        Dim Expected As String
        Dim Timeout As Integer
        Dim console_log_index As Integer
        Dim console_log_usb_inserted As Integer
        Dim AllStringsFound As Boolean
        Dim Timedout As Boolean
        Dim startTime As DateTime
        Dim console_log As ArrayList
        Dim StringFound As Boolean
        Dim T As New TM1
        Dim di As DirectoryInfo
        Dim ExpectedUsbFiles() As String = {"CONFIG.XML", "console_log", "EVENTLOG.XML", "RECORDS.DAT", "RECORDS.XML"}
        Dim ExpectedUsbFile As String
        Dim ReadyToRemove As Boolean = False
        Dim console_log_created As Boolean

        'DR = MessageBox.Show("Install the test USB thumb drive on the host PC", "INSTALL THUMB DRIVE IN PC?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
        '    Return False
        'End If
        If Not YesNo("Install the test USB thumb drive on the host PC", "INSTALL THUMB DRIVE IN PC?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
            Return False
        End If
        CommonLib.Delay(5)
        If Not U.FindThumbDrive(USB_DRIVE) Then
            Form1.AppendText(U.ErrorMsg)
            Return False
        End If
        Form1.AppendText("Found thumb drive at " + USB_DRIVE + ":")

        If Directory.Exists(USB_DRIVE + TM1_SN) Then
            Form1.AppendText("Removing directory " + USB_DRIVE + TM1_SN + " from thumb drive")
            Try
                Directory.Delete(USB_DRIVE + TM1_SN, True)
            Catch ex As Exception
                Form1.AppendText("Problem removing:  " + USB_DRIVE + TM1_SN)
                Form1.AppendText(ex.ToString)
                Return False
            End Try
        End If
        If Exists(USB_DRIVE + "sensorfw.hex") Then
            Try
                File.Delete(USB_DRIVE + "sensorfw.hex")
            Catch ex As Exception
                Form1.AppendText("Problem removing:  " + USB_DRIVE + "sensorfw.hex")
                Return False
            End Try
        End If
        If Exists(USB_DRIVE + "tm1fw.hex") Then
            Try
                File.Delete(USB_DRIVE + "tm1fw.hex")
            Catch ex As Exception
                Form1.AppendText("Problem removing:  " + USB_DRIVE + "tm1fw.hex")
                Return False
            End Try
        End If
        'CommonLib.Delay(8)

        'If Not U.EjectUsbDrive(USB_DRIVE) Then
        '    Form1.AppendText(U.ErrorMsg)
        '    Return False
        'End If
        Try
            Directory.CreateDirectory(USB_DRIVE + "autorun.inf")
        Catch ex As Exception
            Form1.AppendText("Problem creating directory 'autorun.inf' on the thumb drive")
            Form1.AppendText(ex.ToString)
            Return False
        End Try
        If Not U.EjectUsbDrive_AutoIt(USB_DRIVE) Then
            Form1.AppendText(U.ErrorMsg)
            Return False
        End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If

        If Not SF.Cmd(SerialPort, Response, "rm console_log", 10) Then
            Form1.AppendText("Problem removing console_log")
            Form1.AppendText(SF.ErrorMsg)
            Return False
        End If

        ' SF.Close()

        'DR = MessageBox.Show("Install the test USB thumb drive in the TM1", "INSTALL THUMB DRIVE IN TM1?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
        '    Return False
        'End If
        If Not YesNo("Install the test USB thumb drive in the TM1", "INSTALL THUMB DRIVE IN TM1?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
            Return False
        End If
        'CommonLib.Delay(10)

        'results = SF.Connect(SerialPort)
        'If Not results.PassFail Then
        'Return False
        'End If

        startTime = Now
        While Not console_log_created And Now.Subtract(startTime).TotalSeconds < 60
            If Not T.ReadConsoleLog(console_log) Then
                If Not T.ErrorMsg.Contains("console_log file does not exist") Then
                    Form1.AppendText(T.ErrorMsg)
                    Return False
                Else
                    CommonLib.Delay(5)
                End If
            Else
                console_log_created = True
            End If
        End While
        If Not console_log_created Then
            Form1.AppendText("Timeout waiting for console_log to be created")
            Return False
        End If


        ExpectedStrings = New ArrayList

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "USB Mass storage device opened")
        ExpectedStringItem.Add("Timeout", 10)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "File System opened")
        ExpectedStringItem.Add("Timeout", 10)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "records.xml exported")
        ExpectedStringItem.Add("Timeout", 600)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "eventlog exported")
        ExpectedStringItem.Add("Timeout", 30)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "config exported")
        ExpectedStringItem.Add("Timeout", 30)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "console log exported")
        ExpectedStringItem.Add("Timeout", 30)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "Exporting Data complete")
        ExpectedStringItem.Add("Timeout", 10)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "Ready for user to remove USB stick")
        ExpectedStringItem.Add("Timeout", 30)
        ExpectedStrings.Add(ExpectedStringItem)

        console_log_index = 0
        console_log_usb_inserted = console_log_index
        AllStringsFound = False
        ExpectedStringIndex = 0
        startTime = Now
        Expected = ExpectedStrings(ExpectedStringIndex)("Expected")
        Timeout = ExpectedStrings(ExpectedStringIndex)("Timeout")
        Timedout = False
        While Not AllStringsFound And Not Timedout
            If Not T.ReadConsoleLog(console_log) Then
                Form1.AppendText(T.ErrorMsg)
                Return False
            End If
            For i = console_log_index To console_log.Count - 1
                Form1.AppendText(console_log(i))
                If console_log(i).contains(Expected) Then
                    ExpectedStringIndex += 1
                    If ExpectedStringIndex = ExpectedStrings.Count Then
                        AllStringsFound = True
                        Exit For
                    End If
                    startTime = Now
                    Expected = ExpectedStrings(ExpectedStringIndex)("Expected")
                    Timeout = ExpectedStrings(ExpectedStringIndex)("Timeout")
                    StringFound = True
                    'Else
                    '    If Now.Subtract(startTime).TotalSeconds > Timeout Then
                    '        Timedout = True
                    '    End If
                End If
                If console_log(i).contains("Ready for user to remove USB stick") Then
                    ReadyToRemove = True
                End If
            Next
            If Not AllStringsFound Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
                If Now.Subtract(startTime).TotalSeconds > Timeout Then
                    Timedout = True
                End If
                If ReadyToRemove Then
                    Timedout = True
                End If
                CommonLib.Delay(5)
            End If
            console_log_index = console_log.Count
        End While
        If Timedout Then
            Form1.AppendText("Timeout waiting for '" + Expected + "'In console log")
            Return False
        End If

        'DR = MessageBox.Show("Remove Thumb drive from TM1 and install on PC", "MOVE THUMB DRIVE FROM TM1 TO PC?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
        '    Return False
        'End If
        If Not YesNo("Remove Thumb drive from TM1 and install on PC", "MOVE THUMB DRIVE FROM TM1 TO PC?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
            Return False
        End If
        CommonLib.Delay(5)
        If Not U.FindThumbDrive(USB_DRIVE) Then
            Form1.AppendText(U.ErrorMsg)
            Return False
        End If
        Form1.AppendText("Found thumb drive at " + USB_DRIVE + ":")
        If Not Directory.Exists(USB_DRIVE + TM1_SN) Then
            Form1.AppendText("Cant access TM1 directory on thumb drive:  " + USB_DRIVE + TM1_SN)
            Return False
        End If
        Form1.AppendText(USB_DRIVE + TM1_SN + "\")
        di = New DirectoryInfo(USB_DRIVE + TM1_SN)
        For Each fi In di.GetFiles()
            Form1.AppendText(fi.Name)
        Next
        For Each ExpectedUsbFile In ExpectedUsbFiles
            If Not Exists(USB_DRIVE + TM1_SN + "\" + ExpectedUsbFile) Then
                Form1.AppendText(USB_DRIVE + TM1_SN + "\" + ExpectedUsbFile + " not found")
                Return False
            End If
        Next

        Return True
    End Function
End Class