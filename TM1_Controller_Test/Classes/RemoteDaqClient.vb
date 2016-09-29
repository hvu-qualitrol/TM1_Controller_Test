Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.IO

Public Class RemoteDaqClient
    Private _ErrorMsg As String
    Private remoteDaqObject As RemoteDaq.RemoteDaq
    Private http As HttpClientChannel = Nothing
    Private Locked As Boolean = False

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function FindDAQs(ByRef daqs As Hashtable) As Boolean
        'Dim http As HttpClientChannel = Nothing
        'http = New HttpClientChannel()
        'ChannelServices.RegisterChannel(http, False)
        Dim remoteDaqObject As RemoteDaq.RemoteDaq = Nothing
        remoteDaqObject = DirectCast(Activator.GetObject(GetType(RemoteDaq.RemoteDaq), "http://localhost:1234/RemoteDaq"), RemoteDaq.RemoteDaq)

        'If Not RemoteDaqConnected Then
        '    Try
        '        RemotingConfiguration.Configure("\Temp\DaqClient.exe.config", False)
        '    Catch ex As Exception
        '        _ErrorMsg = "Problem connecting to DAQ Server"
        '        Return False
        '    End Try
        '    RemoteDaqConnected = True
        'End If
        ' Dim remoteDaqObject As New RemoteDaq.RemoteDaq()
        'remoteDaqObject = New RemoteDaq.RemoteDaq()
        Dim StatusStr As String = ""
        If Not remoteDaqObject.FindDAQs(daqs, StatusStr) Then
            _ErrorMsg = StatusStr
            'ChannelServices.UnregisterChannel(http)
            Return False
        End If
        '
        '
        'ChannelServices.UnregisterChannel(http)
        Return True
    End Function

    Public Function InitializeDAQ() As Boolean
        Dim BoardNum As Integer

        BoardNum = USB1208LS_Devices(Form1.USB1208LS_ComboBox.Text)
        'If Not RemoteDaqConnected Then
        '    Try
        '        RemotingConfiguration.Configure("\Temp\DaqClient.exe.config", False)
        '    Catch ex As Exception
        '        _ErrorMsg = "Problem connecting to DAQ Server" + vbCr + vbCr
        '        _ErrorMsg += ex.ToString
        '        Return False
        '    End Try
        'End If
        'remoteDaqObject = New RemoteDaq.RemoteDaq()
        remoteDaqObject = Nothing
        remoteDaqObject = DirectCast(Activator.GetObject(GetType(RemoteDaq.RemoteDaq), "http://localhost:1234/RemoteDaq"), RemoteDaq.RemoteDaq)
        If Not remoteDaqObject.InitializeDAQ(BoardNum) Then
            _ErrorMsg = remoteDaqObject.ErrorMsg
            Return False
        End If
        Return True
    End Function

    Public Function ConfigureAnalogInputs(ByVal config As DaqAnalogConfiguration) As Boolean
        If Not remoteDaqObject.ConfigureAnalogInputs(config) Then
            _ErrorMsg = remoteDaqObject.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function DrivePin(ByVal Bit As Integer, ByVal Value As Integer) As Boolean
        If Not remoteDaqObject.DrivePin(Bit, Value) Then
            _ErrorMsg = remoteDaqObject.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function ReadBit(ByVal Bit As Integer, ByRef Value As Integer) As Boolean
        If Not remoteDaqObject.ReadBit(Bit, Value) Then
            _ErrorMsg = remoteDaqObject.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function GetAnalogReading(ByVal Channel As Integer, ByRef Value As Double) As Boolean
        If Not remoteDaqObject.GetAnalogReading(Channel, Value) Then
            _ErrorMsg = remoteDaqObject.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function LockDaq() As Boolean
        Dim startTime As DateTime = Now
        Dim reportTime As DateTime = Now
        ' Dim Locked As Boolean = False

        DaqLockFile = New FileStream(DaqLockFilename, FileMode.Create, FileAccess.Write, FileShare.Write)

        While Not Locked And Now.Subtract(startTime).TotalSeconds < 120
            Application.DoEvents()
            Try
                DaqLockFile.Lock(0, 100)
                Locked = True
            Catch ex As System.IO.IOException
                If Now.Subtract(reportTime).TotalSeconds > 10 Then
                    Form1.AppendText("DAQ locked by another process")
                    reportTime = Now
                End If
                CommonLib.Delay(1)
                'System.Threading.Thread.Sleep(1000)
            Catch ex As Exception
                _ErrorMsg = ex.ToString
                Return False
            End Try

        End While
        If Not Locked Then
            _ErrorMsg = "Timeout obtaining lock on DAQ"
            Return False
        End If

        If Not StartDaqServer() Then
            Return False
        End If

        http = New HttpClientChannel()
        ChannelServices.RegisterChannel(http, False)

        Return True
    End Function

    Public Function UnlockDaq() As Boolean
        If Not Locked Then
            Return True
        End If

        Try
            DaqLockFile.Unlock(0, 100)
        Catch ex As Exception
            _ErrorMsg = "Probleming unlocking daq" + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        If Not StopDaqServer() Then
            Return False
        End If
        ChannelServices.UnregisterChannel(http)

        Return True
    End Function

    Private Function StartDaqServer() As Boolean
        DaqServerProcess = New Process

        'TODO:  Check if already running and kill
        'TODO:  save pid of when starting
        'TODO:  error msg return from DaqServerProcess
        DaqServerProcess.StartInfo.UseShellExecute = False
        'DaqServerProcess.StartInfo.UseShellExecute = True
        'DaqServerProcess.StartInfo.RedirectStandardOutput = True
        'DaqServerProcess.StartInfo.RedirectStandardError = True
        DaqServerProcess.StartInfo.RedirectStandardOutput = False
        DaqServerProcess.StartInfo.RedirectStandardError = False
        DaqServerProcess.StartInfo.CreateNoWindow = True
        'DaqServerProcess.StartInfo.CreateNoWindow = False
        DaqServerProcess.StartInfo.FileName = UtilitiesDirectory + "RemoteDaqServer.exe"

        Try
            DaqServerProcess.Start()
        Catch ex As Exception
            _ErrorMsg = "Problem starting daq server process" + vbCr + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        If DaqServerProcess.HasExited Then
            _ErrorMsg = "Daq server process exited prematurely" + vbCr + vbCr
            Return False
        End If

        Return True
    End Function

    Private Function StopDaqServer() As Boolean
        Try
            DaqServerProcess.Kill()
        Catch ex As Exception
            _ErrorMsg = "Problem stopping DAQ Server" + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        Return True
    End Function
End Class
