Imports FTD2XX_NET
Imports System.IO

Public Class FT232R
    Private _failure_message As String

    Public Property failure_message As String
        Get
            Return _failure_message
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function Rescan() As Boolean
        Dim ftdi_device As New FTDI
        Dim ftstatus As FTDI.FT_STATUS

        ftstatus = ftdi_device.Rescan()
        If (Not ftstatus = FTDI.FT_STATUS.FT_OK) Then
            _failure_message = ftstatus.ToString
            Return False
        End If

        Return True
    End Function


    Function FtdiDeviceCount(ByRef DeviceCount As UInteger) As Boolean
        Dim ftdi_device As New FTDI
        Dim ftstatus As FTDI.FT_STATUS

        _failure_message = ""
        Try
            ftstatus = ftdi_device.GetNumberOfDevices(DeviceCount)
        Catch ex As Exception
            _failure_message = "Exception getting FTDI device count:  " + ex.ToString
            Return False
        End Try
        If (ftstatus = FTDI.FT_STATUS.FT_OK) Then
            Return True
        Else
            _failure_message = "Error getting FTDI device count: " + ftstatus.ToString
            Return False
        End If
    End Function

    ' Returns the location of ID of FT232R device.    
    Function GetLocation(ByRef Location As UInteger) As Boolean
        Dim DeviceCount As UInteger
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim FoundFT232R As Boolean = False
        Dim retryCnt As Integer = 0

        _failure_message = ""
        If Not FtdiDeviceCount(DeviceCount) Then
            _failure_message = "FtdiDeviceCount() failed"
            Return False
        End If
        If Not (DeviceCount > 0) Then
            _failure_message = "No FTDI devices found"
            Return False
        End If

        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE
        While Not FoundFT232R And retryCnt < 4
            retryCnt += 1
            ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
            If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                _failure_message = "Problem getting FTDI device info"
                Return False
            End If

            For i = 0 To DeviceCount - 1
                If ftdiDeviceList(i).Type = FTDI.FT_DEVICE.FT_DEVICE_232R Then
                    Form1.AppendText("device " + i.ToString + "location id " + ftdiDeviceList(i).LocId.ToString)
                    ' If ftdiDeviceList(i).LocId < &H100 Then
                    If ftdiDeviceList(i).LocId < &H1000 Then
                        If FoundFT232R Then
                            _failure_message = "Multiple FT232R devices found with a single port, expected 1"
                            Return False
                        End If
                        If ftdiDeviceList(i).LocId > 0 Then
                            Location = ftdiDeviceList(i).LocId
                            FoundFT232R = True
                        End If
                    Else
                        Form1.AppendText("Skipping " + ftdiDeviceList(i).LocId.ToString)
                    End If
                End If
            Next
            If Not FoundFT232R Then CommonLib.Delay(5)
        End While
        'ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
        'If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
        '    _failure_message = "Problem getting FTDI device info"
        '    Return False
        'End If

        'For i = 0 To DeviceCount - 1
        '    If ftdiDeviceList(i).Type = FTDI.FT_DEVICE.FT_DEVICE_232R Then
        '        Form1.AppendText("device " + i.ToString + "location id " + ftdiDeviceList(i).LocId.ToString)
        '        ' If ftdiDeviceList(i).LocId < &H100 Then
        '        If ftdiDeviceList(i).LocId < &H1000 Then
        '            If FoundFT232R Then
        '                _failure_message = "Multiple FT232R devices found with a single port, expected 1"
        '                Return False
        '            End If
        '            Location = ftdiDeviceList(i).LocId
        '            FoundFT232R = True
        '        Else
        '            Form1.AppendText("Skipping " + ftdiDeviceList(i).LocId.ToString)
        '        End If
        '    End If
        'Next
        Return FoundFT232R
    End Function

    Public Function CycleComport(ByVal comport As String) As Boolean
        Dim DeviceCount As Integer
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim ComportForDevice As String
        Dim DeviceCycled As Boolean = False

        _failure_message = 0
        If Not FtdiDeviceCount(DeviceCount) Then
            _failure_message = "Problem getting FTDI Device Count"
            Return False
        End If
        If Not DeviceCount > 0 Then
            _failure_message = "No FTDI devices found"
            Return False
        End If

        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE
        ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
            Return False
        End If
        For i = 0 To DeviceCount - 1
            Try
                ftstatus = ftdi_device.OpenByLocation(ftdiDeviceList(i).LocId)
            Catch ex As Exception
                Continue For
            End Try
            If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                Try
                    ftdi_device.Close()
                Catch ex As Exception
                End Try
                Continue For
            End If
            ftstatus = ftdi_device.GetCOMPort(ComportForDevice)
            If ComportForDevice = comport Then
                Try
                    ftstatus = ftdi_device.CyclePort()
                Catch ex As Exception
                    _failure_message = "Exception cycling " + comport + vbCr + ex.ToString
                    ftdi_device.Close()
                    Return False
                End Try
                If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                    _failure_message = "Problem cycling " + comport + ", ftstatus = " + ftstatus.ToString
                    ftdi_device.Close()
                    Return False
                End If
                DeviceCycled = True
            End If
            ftdi_device.Close()
        Next

        Return DeviceCycled
    End Function


    Public Function FindLocationForSN(ByVal SN As String, ByRef LocID As Integer) As Boolean
        Dim DeviceCount As Integer
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim found_SN As Boolean = False
        Dim retryCnt As Integer = 0

        If Not FtdiDeviceCount(DeviceCount) Then
            Return False
        End If
        If Not DeviceCount > 0 Then
            Return False
        End If

        _failure_message = ""
        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE
        While Not found_SN And retryCnt < 4
            retryCnt += 1
            ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
            If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
                Return False
            End If
            For i = 0 To DeviceCount - 1
                Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId))
                Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr)
                Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
                If ftdiDeviceList(i).SerialNumber = SN Then
                    Form1.AppendText("Location found for " + SN)
                    LocID = ftdiDeviceList(i).LocId
                    Return True
                End If
            Next
            CommonLib.Delay(5)
        End While

        'ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
        'If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
        '    _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
        '    Return False
        'End If
        'For i = 0 To DeviceCount - 1
        '    Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId))
        '    Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr)
        '    Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
        '    If ftdiDeviceList(i).SerialNumber = SN Then
        '        Form1.AppendText("Location found for " + SN)
        '        LocID = ftdiDeviceList(i).LocId
        '        Return True
        '    End If
        'Next
        _failure_message = "Did not find FTDI device with SN " + SN

        Return False
    End Function

    ''' <summary>
    ''' Here we are trying to connec to the controller boards USB port.
    ''' </summary>
    ''' <param name="SN">Serial number of the TM1</param>
    ''' <param name="ComPort">Comm port used by the Measurement Computering Test Board</param>
    ''' <param name="FtdiDevices">Hashtable to hold the entries for all of the FTDI devices found.</param>
    ''' <returns>True means success</returns>
    ''' <remarks></remarks>
    Public Function FindComportForSN(ByVal SN As String, ByRef ComPort As String, Optional ByRef FtdiDevices As Hashtable = Nothing, Optional ByVal newFetch As Boolean = False) As Boolean
        Dim DeviceCount As Integer
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim found_blank_SN As Boolean = False
        Dim retryCnt As Integer = 0
        Dim found_SN As Boolean = False
        'Dim FtdiHash As New Hashtable
        Dim ComportForDevice As String = ""
        Dim FtdiDevice As Hashtable

        FtdiDevices = New Hashtable

        ComPort = "UNKNOWN"

        If Not newFetch Then
            If FtdiHash.Contains(SN) Then
                If System.IO.Ports.SerialPort.GetPortNames().Contains(FtdiHash(SN)) Then
                    ComPort = FtdiHash(SN)
                    Form1.AppendText(SN + " mapped to " + ComPort + ", found from hash")
                    Return True
                Else
                    FtdiHash.Remove(SN)
                End If
            End If

            GetFtdiMappingsFromMemory()
            If FtdiHash.Contains(SN) Then
                If System.IO.Ports.SerialPort.GetPortNames().Contains(FtdiHash(SN)) Then
                    ComPort = FtdiHash(SN)
                    Form1.AppendText(SN + " mapped to " + ComPort + ", found from memory")
                    Return True
                Else
                    FtdiHash.Remove(SN)
                End If
            End If
        End If

        ' Uses a reference to set the value
        ' DeviceCount to the number of devices
        ' found. This value is gathered by the 
        ' FTDI library.
        If Not FtdiDeviceCount(DeviceCount) Then
            Return False
        End If
        If Not DeviceCount > 0 Then
            Return False
        End If

        _failure_message = ""
        Dim startTime As DateTime = Now
        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE

        found_blank_SN = True
        While (Not found_SN And found_blank_SN And retryCnt < 3)
            ' The list is a Node object which is part of the 
            ' FTDI library. It is passed as a reference to the
            ' GetDeviceList function and populated with an array
            ' of devices found by the library.
            Try
                ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
                If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                    _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
                    Return False
                End If

                found_blank_SN = False
                If retryCnt > 0 Then
                    CommonLib.Delay(15)
                    Form1.AppendText("RETRY")
                End If
                retryCnt += 1

                ' Here we walk the number of devices found and try to find the TM1 on the other side.
                For i = 0 To DeviceCount - 1
                    Try
                        ftstatus = ftdi_device.OpenByLocation(ftdiDeviceList(i).LocId)
                    Catch ex As Exception
                        Exit For
                    End Try
                    Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId))
                    Form1.AppendText("SN=" + ftdiDeviceList(i).SerialNumber)
                    Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
                    ftstatus = ftdi_device.GetCOMPort(ComportForDevice)
                    If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                        Form1.AppendText("GetCOMPort failed")
                        found_blank_SN = True
                        Continue For
                    End If
                    Form1.AppendText("COM PORT = " + ComportForDevice)
                    'ftdi_device.CyclePort()
                    'ftdi_device.Close()
                    If ftdiDeviceList(i).SerialNumber.Length = 0 Then
                        found_blank_SN = True
                        Continue For
                    End If
                    If ftdiDeviceList(i).SerialNumber = SN And ComportForDevice.Length <> 0 Then
                        found_SN = True
                        ComPort = ComportForDevice
                        ftdi_device.CyclePort()
                    End If
                    ftdi_device.Close()
                    If Not FtdiHash.Contains(ftdiDeviceList(i).SerialNumber) Then
                        FtdiHash.Add(ftdiDeviceList(i).SerialNumber, ComportForDevice)
                    End If
                    FtdiDevice = New Hashtable
                    If Not FtdiDevices.Contains(ftdiDeviceList(i).SerialNumber) Then
                        FtdiDevice.Add("COM", ComportForDevice)
                        FtdiDevice.Add("LOCID", ftdiDeviceList(i).LocId)
                        FtdiDevices.Add(ftdiDeviceList(i).SerialNumber, FtdiDevice)
                    End If
                Next

                'For i = 0 To DeviceCount - 1
                '    Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId), True)
                '    Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr, True)
                '    Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
                '    If ftdiDeviceList(i).SerialNumber = "" Then
                '        Form1.AppendText("SN not programmed in this device?")
                '        found_blank_SN = True
                '    End If
                '    If ftdiDeviceList(i).SerialNumber = SN Then
                '        found_SN = True
                '        Try
                '            ftstatus = ftdi_device.OpenBySerialNumber(ftdiDeviceList(i).SerialNumber)
                '        Catch ex As Exception
                '            _failure_message = "Exception opening FTDI device:  " + ex.ToString
                '            Return False
                '        End Try
                '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                '            _failure_message = "Error opening FTDI device:  " + ftstatus.ToString
                '            Return False
                '        End If
                '        ftstatus = ftdi_device.GetCOMPort(ComPort)
                '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                '            _failure_message = "Error getting FTDI comport:  " + ftstatus.ToString
                '            Return False
                '        End If
                '        ftstatus = ftdi_device.Close()
                '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                '            _failure_message = "Error closing FTDI device  " + ftstatus.ToString
                '            Return False
                '        End If
                '        Form1.AppendText("resetting " + ftdiDeviceList(i).SerialNumber)
                '        ftdi_device.ResetPort()

                '        For j = 0 To DeviceCount - 1
                '            'If j = i Then Continue For
                '            If Not j = i And Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Then
                '                Continue For
                '            End If
                '            Form1.AppendText("cycling " + ftdiDeviceList(j).SerialNumber)
                '            ftdi_device.OpenByIndex(j)
                '            ftdi_device.CyclePort()
                '            ftdi_device.Close()
                '        Next
                '        Return True
                '    End If
                'Next
            Catch ex As Exception
                Form1.AppendText("FindComportForSN() Caught " + ex.ToString)
                Return False
            End Try
        End While

        'For i = 0 To DeviceCount - 1
        '    Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId), True)
        '    Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr, True)
        '    Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
        '    If ftdiDeviceList(i).SerialNumber = "" Then
        '        found_blank_SN = True
        '    End If
        '    If ftdiDeviceList(i).SerialNumber = SN Then
        '        found_SN = True
        '        Try
        '            ftstatus = ftdi_device.OpenBySerialNumber(ftdiDeviceList(i).SerialNumber)
        '        Catch ex As Exception
        '            _failure_message = "Exception opening FTDI device:  " + ex.ToString
        '            Return False
        '        End Try
        '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
        '            _failure_message = "Error opening FTDI device:  " + ftstatus.ToString
        '            Return False
        '        End If
        '        ftstatus = ftdi_device.GetCOMPort(ComPort)
        '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
        '            _failure_message = "Error getting FTDI comport:  " + ftstatus.ToString
        '            Return False
        '        End If
        '        ftstatus = ftdi_device.Close()
        '        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
        '            _failure_message = "Error closing FTDI device  " + ftstatus.ToString
        '            Return False
        '        End If
        '        Form1.AppendText("resetting " + ftdiDeviceList(i).SerialNumber)
        '        ftdi_device.ResetPort()


        '        For j = 0 To DeviceCount - 1
        '            'If j = i Then Continue For
        '            If Not j = i And Form1.AssemblyLevel_ComboBox.Text = "SUB_ASSEMBLY" Then
        '                Continue For
        '            End If
        '            Form1.AppendText("cycling " + ftdiDeviceList(j).SerialNumber)
        '            ftdi_device.OpenByIndex(j)
        '            ftdi_device.CyclePort()
        '            ftdi_device.Close()
        '        Next
        '        Return True
        '    End If
        'Next


        ' Return False
        ShareFtdiMappings()
        GetFtdiMappingsFromMemory()
        Return found_SN
    End Function

    Function SetSN(ByVal ftdi_device As FTDI, SN As String) As Boolean
        Dim myEEData As FTDI.FT232R_EEPROM_STRUCTURE
        Dim ftstatus As FTDI.FT_STATUS

        _failure_message = ""
        If Not ftdi_device.IsOpen Then
            _failure_message = "FTDI device is not open"
            Return False
        End If

        myEEData = New FTDI.FT232R_EEPROM_STRUCTURE()
        Try
            ftstatus = ftdi_device.ReadFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception reading FTDI eeprom:  " + ex.ToString
            Return False
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error reading FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        myEEData.SerialNumber = SN
        Try
            ftstatus = ftdi_device.WriteFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception writing FTDI eeprom:  " + ex.ToString
            Return (False)
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error writing FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        Return True
    End Function

    Function SetDescription(ByVal ftdi_device As FTDI, description As String) As Boolean
        Dim myEEData As FTDI.FT232R_EEPROM_STRUCTURE
        Dim ftstatus As FTDI.FT_STATUS

        _failure_message = ""
        If Not ftdi_device.IsOpen Then
            _failure_message = "FTDI device is not open"
            Return False
        End If

        myEEData = New FTDI.FT232R_EEPROM_STRUCTURE()
        Try
            ftstatus = ftdi_device.ReadFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception reading FTDI eeprom:  " + ex.ToString
            Return False
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error reading FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        myEEData.Description = description
        Try
            ftstatus = ftdi_device.WriteFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception writing FTDI eeprom:  " + ex.ToString
            Return (False)
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error writing FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        Return True
    End Function

    Function CycleDevice(ByVal ftdi_device As FTDI, LocId As UInteger) As Boolean
        Dim ftstatus As FTDI.FT_STATUS
        Dim startTime As DateTime = Now

        _failure_message = ""
        Try
            ftstatus = ftdi_device.CyclePort()
        Catch ex As Exception
            _failure_message = "Exception cycling FTDI device:  " + ex.ToString
            Return False
        End Try
        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error cycling FTDI device:  " + ftstatus.ToString
            Return False
        End If
        Do
            ftstatus = ftdi_device.OpenByLocation(LocId)
            System.Threading.Thread.Sleep(1000)
        Loop Until (ftstatus = FTDI.FT_STATUS.FT_OK) Or (Now.Subtract(startTime).TotalSeconds > 30)

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Timeout resetting FTDI device "
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Tries to lock the Ftdi Serial to USB converter.
    ''' </summary>
    ''' <returns>A success is true</returns>
    ''' <remarks></remarks>
    Public Function LockFtdi() As Boolean
        Dim startTime As DateTime = Now
        Dim Locked As Boolean = False
        Dim MsgCnt As Integer = 0

        FtdiLockFile = New FileStream(FtdiLockFilename, FileMode.Create, FileAccess.Write, FileShare.Write)

        While Not Locked And Now.Subtract(startTime).TotalSeconds < 120
            Application.DoEvents()
            Try
                FtdiLockFile.Lock(0, 100)
                ' DaqLockFile.Lock(0, 100)
                Locked = True
            Catch ex As System.IO.IOException
                If MsgCnt Mod 15 = 0 Then
                    Form1.AppendText("locked by another process, waiting to obtain lock")
                End If
                System.Threading.Thread.Sleep(1000)
                MsgCnt += 1
            Catch ex As Exception
                _failure_message = ex.ToString
                Return False
            End Try

        End While
        If Not Locked Then
            _failure_message = "Timeout obtaining lock on FTDI"
        End If

        Return Locked
    End Function

    Public Function UnlockFtdi() As Boolean
        Try
            FtdiLockFile.Unlock(0, 100)
            ' DaqLockFile.Unlock(0, 100)
            Return True
        Catch ex As Exception
            _failure_message = ex.ToString
        End Try

        Return False
    End Function

    Public Function ShareFtdiMappings()
        Dim smm As New SharedMemoryMessaging
        Dim msg As String = ""

        For Each SN In FtdiHash.Keys
            'If msg.Length > 0 Then
            '    msg += ":"
            'End If
            msg += SN + "," + FtdiHash(SN) + ":"
        Next
        If Not smm.SendMessage(msg) Then
            Form1.AppendText("smm failed")
            Return False
        End If

        Return True
    End Function

    Private Function GetFtdiMappingsFromMemory() As Boolean
        Dim smm As New SharedMemoryMessaging
        Dim msg As String
        Dim Fields() As String

        FtdiHash = New Hashtable

        If Not smm.ReadMessage(msg) Then
            Return False
        End If
        For Each Line In Split(msg, ":")
            Fields = Split(Line, ",")
            If Fields.Length = 2 Then
                Fields(1) = Fields(1).Trim
                If Not FtdiHash.Contains(Fields(0)) Then
                    FtdiHash.Add(Fields(0), Fields(1))
                End If
            End If
        Next
        Return True
    End Function
End Class
'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           