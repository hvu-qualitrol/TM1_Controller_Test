Imports System.Text.RegularExpressions
Imports System.IO.File
Imports System.IO

Partial Class Tests
    Public Shared Function Test_H2SCAN_UPDATE() As Boolean
        Dim Version As String
        Dim MatchedVersion As String
        Dim VersionWord As String
        Dim Checksum As String
        Dim T As New TM1
        Dim console_log As ArrayList
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        'Dim DR As DialogResult
        Dim console_log_index As Integer
        Dim console_log_usb_inserted As Integer
        Dim startTime As DateTime
        Dim Response As String
        Dim ExpectedStrings As ArrayList
        Dim ExpectedStringItem As Hashtable
        Dim ExpectedStringIndex As Integer
        Dim Expected As String
        Dim Timeout As Integer
        Dim StringFound As Boolean
        Dim AllStringsFound As Boolean
        Dim Timedout As Boolean
        Dim ReportedMatchVersion As String
        Dim ReportedUpgradeVersion As String
        Dim UpgradeSuccessful As Boolean
        Dim ExpectedUpdateStrings(5) As String
        Dim USB_DRIVE As String
        Dim fw_file As String
        Dim dest_file As String
        Dim U As New USB
        Dim ExpectedUsbFiles() As String = {"CONFIG.XML", "console_log", "EVENTLOG.XML", "RECORDS.DAT", "RECORDS.XML"}
        Dim ExpectedUsbFile As String
        Dim di As DirectoryInfo
        Dim console_log_exists As Boolean

        Version = Products(Product)("sensor firmware version")
        MatchedVersion = Products(Product)("sensor firmware matched version")
        VersionWord = Regex.Replace(Version, "\.", "")
        VersionWord = Regex.Split(VersionWord, "(\d+)")(1)
        Checksum = Products(Product)("sensor firmware checksum")

        If Not YesNo("Install the test USB thumb drive on the host PC", "INSTALL THUMB DRIVE IN PC?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
            Return False
        End If
        'DR = MessageBox.Show("Install the test USB thumb drive on the host PC", Form1.Serial_Number_Entry.Text + ":  INSTALL THUMB DRIVE IN PC?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
        '    Return False
        'End If
        CommonLib.Delay(7)
        If Not U.FindThumbDrive(USB_DRIVE) Then
            Form1.AppendText(U.ErrorMsg)
            Return False
        End If
        Form1.AppendText("Found thumb drive at " + USB_DRIVE + ":")
        fw_file = BaseFWDirectory + "Serv" + VersionWord + "." + Checksum + ".hex"
        If Not Exists(fw_file) Then
            Form1.AppendText("sensor fw file " + fw_file + " not found")
            Return False
        End If
        dest_file = USB_DRIVE + "sensorfw.hex"
        Form1.AppendText("copy " + fw_file + " " + dest_file)
        Try
            FileIO.FileSystem.CopyFile(fw_file, dest_file, True)
        Catch ex As Exception
            Form1.AppendText("Problem copying sensor fw file")
            Form1.AppendText(ex.ToString)
            Return False
        End Try

        If Directory.Exists(USB_DRIVE + Form1.Serial_Number) Then
            Try
                Directory.Delete(USB_DRIVE + Form1.Serial_Number, True)
            Catch ex As Exception
                Form1.AppendText("Problem removing:  " + USB_DRIVE + Form1.Serial_Number)
                Form1.AppendText(ex.ToString)
                Return False
            End Try
        End If

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

        ' Check to see whether H2Scan needs to update its firmware so the test can act accordingly
        Dim VersionInfo As Hashtable = Nothing
        Dim updateH2ScanFirmware As Boolean = False
        If T.GetVersionInfo(VersionInfo) Then
            If VersionInfo("sensor firmware version") <> Products(Product)("sensor firmware version") Then
                updateH2ScanFirmware = True
            End If
        End If

        SF.Close()

        'DR = MessageBox.Show("Install the test USB thumb drive on the TM1", Form1.Serial_Number_Entry.Text + ":  INSTALL THUMB DRIVE IN TM1?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
        '    Return False
        'End If
        If Not YesNo("Install the test USB thumb drive on the TM1", "INSTALL THUMB DRIVE IN TM1?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in TM1.")
            Return False
        End If

        If updateH2ScanFirmware Then
            Form1.AppendText("Updating H2Scan firmware...")
            WaitTimeout(600.0, "Updatig H2Scan firmware... ")
        End If

        results = SF.Connect(SerialPort)
        If Not results.PassFail Then
            Return False
        End If
        Form1.AppendText("Successfully reconnected!")

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
        ExpectedStringItem.Add("Timeout", 120)
        ExpectedStrings.Add(ExpectedStringItem)

        ExpectedStringItem = New Hashtable
        ExpectedStringItem.Add("Expected", "Ready for user to remove USB stick")
        ExpectedStringItem.Add("Timeout", 30)
        ExpectedStrings.Add(ExpectedStringItem)

        Form1.AppendText("Start trying to read console log file...")
        startTime = Now
        console_log_exists = False
        While (Not console_log_exists And Now.Subtract(startTime).TotalSeconds < 60)
            If Not T.ReadConsoleLog(console_log) Then
                If Not T.ErrorMsg.Contains("file does not exist") Then
                    Form1.AppendText(T.ErrorMsg)
                    Return False
                End If
            Else
                console_log_exists = True
            End If
        End While
        If Not console_log_exists Then
            Form1.AppendText("Timeout waiting for console_log file to be generated")
            Return False
        End If

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
                Else
                    If Now.Subtract(startTime).TotalSeconds > Timeout Then
                        Timedout = True
                    End If
                End If
            Next
            If Not AllStringsFound Then
                Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(startTime).TotalSeconds).ToString
                CommonLib.Delay(5)
            End If
            console_log_index = console_log.Count
        End While
        If Timedout Then
            Form1.AppendText("Timeout waiting for '" + Expected + "'In console log")
            Return False
        End If

        ' H2scan only reports the first three letters of the version
        VersionWord = VersionWord.Substring(0, 3)

        ' Check if sensor updated or was already current
        ReportedMatchVersion = "NONE"
        ReportedUpgradeVersion = "NONE"
        UpgradeSuccessful = False
        For i = console_log_usb_inserted To console_log.Count - 1
            If Regex.IsMatch(console_log(i), "Sensor firmware upgrade version matches sensor \(\d+\)") Then
                ReportedMatchVersion = Regex.Split(console_log(i), "Sensor firmware upgrade version matches sensor \((\d+)\)")(1)
            End If
            If Regex.IsMatch(console_log(i), "Changing Sensor firmware \(\d+ to \d+\)") Then
                ReportedUpgradeVersion = Regex.Split(console_log(i), "Changing Sensor firmware \(\d+ to (\d+)\)")(1)
            End If
            If console_log(i).Contains("Firmware change complete") Then
                UpgradeSuccessful = True
            End If
        Next
        If ReportedMatchVersion = "NONE" Then
            If Not UpgradeSuccessful Then
                Form1.AppendText("H2SCAN sensor update not successful")
                Return False
            Else
                If Not ReportedUpgradeVersion = MatchedVersion Then
                    Form1.AppendText("Upgrade version = " + ReportedUpgradeVersion + ", Expected " + MatchedVersion)
                    Return False
                End If
            End If
        Else
            If UpgradeSuccessful Then
                Form1.AppendText("Unexpected H2SCAN reflash detected")
                Return False
            ElseIf Not ReportedUpgradeVersion = "NONE" Then
                Form1.AppendText("Unexpected H2SCAN reflash detected")
                Return False
            End If
            If Not ReportedMatchVersion = MatchedVersion Then
                Form1.AppendText("Reported FW version=" + ReportedMatchVersion + ", Expected=" + MatchedVersion)
                Return False
            End If
        End If

        'DR = MessageBox.Show("Remove Thumb drive from TM1 and install on PC", Form1.Serial_Number_Entry.Text + ":  MOVE THUMB DRIVE FROM TM1 TO PC?", MessageBoxButtons.YesNo)
        'If DR = DialogResult.No Then
        '    Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
        '    Return False
        'End If
        If Not YesNo("Remove Thumb drive from TM1 and install on PC", "MOVE THUMB DRIVE FROM TM1 TO PC?") Then
            Form1.AppendText("Operator indicated problem installing the USB thumb drive in PC.")
            Return False
        End If
        CommonLib.Delay(7)
        If Not U.FindThumbDrive(USB_DRIVE) Then
            Form1.AppendText(U.ErrorMsg)
            Return False
        End If
        Form1.AppendText("Found thumb drive at " + USB_DRIVE + ":")
        If Not Directory.Exists(USB_DRIVE + Form1.Serial_Number) Then
            Form1.AppendText("Cant access TM1 directory on thumb drive:  " + USB_DRIVE + Form1.Serial_Number)
            Return False
        End If
        Form1.AppendText(USB_DRIVE + Form1.Serial_Number + "\")
        di = New DirectoryInfo(USB_DRIVE + Form1.Serial_Number)
        For Each fi In di.GetFiles()
            Form1.AppendText(fi.Name)
        Next
        For Each ExpectedUsbFile In ExpectedUsbFiles
            If Not Exists(USB_DRIVE + Form1.Serial_Number + "\" + ExpectedUsbFile) Then
                Form1.AppendText(USB_DRIVE + Form1.Serial_Number + "\" + ExpectedUsbFile + " not found")
                Return False
            End If
        Next



        Return True
    End Function
End Class