Imports System.IO

Public Class DAQ_functions
    Private ULStat As MccDaq.ErrorInfo
    Private ErrorMessage As String
    Private _ErrorMsg As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function GetErrorMsg() As String
        Return ErrorMessage + vbCr + ULStat.Message
    End Function

    Public Function FindDAQs(ByRef daqs As Hashtable) As Boolean
        Dim NumBoards As Integer = 4
        Dim DaqBoard As MccDaq.MccBoard
        Dim ULStat As MccDaq.ErrorInfo
        Dim ConfigVal As Integer
        Dim SN As String
        Dim Success As Boolean = False
        Dim NumPorts As Integer

        daqs = New Hashtable

        If Not LockDaq() Then
            Return False
        End If

        For i = 0 To NumBoards - 1
            Application.DoEvents()
            DaqBoard = New MccDaq.MccBoard(i)
            ULStat = DaqBoard.BoardConfig.GetBoardType(ConfigVal)
            If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
                _ErrorMsg = ULStat.Message
                Success = False
                Exit For
            End If
            ULStat = DaqBoard.BoardConfig.GetDiNumDevs(NumPorts)
            If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
                Continue For
            End If
            If Not NumPorts = 2 Then
                Continue For
            End If
            If ConfigVal > 0 Then
                DaqBoard.GetConfigString(2, 0, 224, SN, 32)
                Form1.AppendText("Found DAQ " + SN + " at board# " + i.ToString)
                If daqs.Contains(SN) Then
                    daqs(SN) = i
                Else
                    daqs.Add(SN, i)
                End If
                Success = True
            End If
            Form1.AppendText("Boardname = " + DaqBoard.BoardName.ToString)
        Next
        If daqs.Count = 0 Then
            _ErrorMsg = "No USB1208-LS daq devices found"
        End If

        If Not UnlockDaq() Then
            Return False
        End If
        Return Success
    End Function

    Private Function LockDaq() As Boolean
        Dim startTime As DateTime = Now
        Dim Locked As Boolean = False

        DaqLockFile = New FileStream(DaqLockFilename, FileMode.Create, FileAccess.Write, FileShare.Write)

        While Not Locked And Now.Subtract(startTime).TotalSeconds < 30
            Application.DoEvents()
            Try
                DaqLockFile.Lock(0, 100)
                Locked = True
            Catch ex As System.IO.IOException
                Form1.AppendText("locked by another process")
                System.Threading.Thread.Sleep(1000)
            Catch ex As Exception
                _ErrorMsg = ex.ToString
                Return False
            End Try

        End While
        If Not Locked Then
            _ErrorMsg = "Timeout obtaining lock on DAQ"
        End If

        Return Locked
    End Function

    Private Function UnlockDaq() As Boolean
        Try
            DaqLockFile.Unlock(0, 100)
            Return True
        Catch ex As Exception
            _ErrorMsg = ex.ToString
        End Try

        Return False
    End Function


    Public Function InitializeDAQ()
        Dim NumPorts As Integer
        Dim BoardNum As Integer

        BoardNum = USB1208LS_Devices(Form1.USB1208LS_ComboBox.Text)

        'Daq = New MccDaq.MccBoard(0)
        Form1.AppendText("Using USB1208LS board number " + BoardNum.ToString)
        Daq = New MccDaq.MccBoard(BoardNum)

        ' Verify MccDaq is accessible
        ULStat = Daq.BoardConfig.GetDiNumDevs(NumPorts)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Problem getting number of devices"
            Return False
        End If
        If Not NumPorts = 2 Then
            ErrorMessage = "Found " + NumPorts.ToString + " ports, expected 2"
            Return False
        End If

        ' Port B configured as an input port
        ULStat = Daq.DConfigPort(MccDaq.DigitalPortType.FirstPortB, MccDaq.DigitalPortDirection.DigitalIn)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Error configuring port B as an input port."
            Return False
        End If

        ' Port A configured as an output port
        ULStat = Daq.DConfigPort(MccDaq.DigitalPortType.FirstPortA, MccDaq.DigitalPortDirection.DigitalOut)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Error configuring port A as an output port."
            Return False
        End If

        Return True
    End Function

    Public Function ConfigureAnalogInputs(ByVal config As DaqAnalogConfiguration)
        Dim NumChans As Integer

        If config = DaqAnalogConfiguration.SE Then
            NumChans = 8
        Else
            NumChans = 4
        End If
        ULStat = Daq.BoardConfig.SetNumAdChans(NumChans)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Error configuring analog inputs"
            Return False
        End If

        Return True
    End Function

    Public Function DrivePin(ByVal Bit As Integer, ByVal Value As Integer) As Boolean
        Dim DriveLevel As MccDaq.DigitalLogicState

        If Value = 1 Then
            DriveLevel = MccDaq.DigitalLogicState.High
        Else
            DriveLevel = MccDaq.DigitalLogicState.Low
        End If
        ULStat = Daq.DBitOut(MccDaq.DigitalPortType.FirstPortA, Bit, DriveLevel)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Problem setting bit=" + Bit.ToString + " to value=" + Value.ToString
            Return False
        End If
        Return True
    End Function

    Public Function ReadBit(ByVal Bit As Integer, ByRef Value As Integer)
        Dim bitValue As MccDaq.DigitalLogicState

        ULStat = Daq.DBitIn(MccDaq.DigitalPortType.FirstPortA, Bit, bitValue)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Problem reading bit " + Bit.ToString
            Return False
        End If
        ' Form1.RTB.AppendText(Bit.ToString + " = " + bitValue.ToString + vbCr)
        If bitValue = MccDaq.DigitalLogicState.High Then
            Value = 1
        Else
            Value = 0
        End If

        Return True
    End Function

    Public Function GetAnalogReading(ByVal Channel As Integer, ByRef Value As Double)
        Dim RegisterReading As Short

        ULStat = Daq.AIn(Channel, MccDaq.Range.Bip10Volts, RegisterReading)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Error reading analog input"
            Return False
        End If

        ULStat = Daq.ToEngUnits32(MccDaq.Range.Bip10Volts, RegisterReading, Value)
        If Not ULStat.Value = MccDaq.ErrorInfo.ErrorCode.NoErrors Then
            ErrorMessage = "Error converting analog register value = " + RegisterReading.ToString + " to double"
            Return False
        End If

        Return True
    End Function

End Class

Public Class DaqBackground
    Public DaqThread As System.Threading.Thread
    Friend CMD As String
    Friend CMD_START As Boolean = False
    Friend CMD_DONE As Boolean = False
    Friend StopDaq As Boolean = False
    Friend RetVal As Boolean
    Friend _ErrorMsg As String = ""

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Function Start() As Boolean
        Dim startTime As DateTime = Now

        DaqThread = New System.Threading.Thread(AddressOf DaqProcess)
        DaqThread.Start()

        Return True
    End Function

    Function Command(ByVal DaqCmd As String) As Boolean
        Dim startTime As DateTime = Now
        CMD = DaqCmd
        CMD_START = True
        While Not CMD_DONE And Now.Subtract(startTime).TotalSeconds < 20
            Application.DoEvents()
        End While

        Return RetVal
    End Function

    Sub DaqProcess()
        Dim D As New DAQ_functions
        Dim CmdSuccess As Boolean

        While Not StopDaq
            If CMD_START Then
                CMD_START = False
                Select Case CMD
                    Case "FIND"
                        If Not D.FindDAQs(USB1208LS_Devices) Then
                            CmdSuccess = False
                            _ErrorMsg = D.ErrorMsg
                        Else
                            CmdSuccess = True
                        End If
                        'CmdSuccess = D.FindDAQs(USB1208LS_Devices)
                        'MsgBox(USB1208LS_Devices.Count.ToString)
                End Select
                CMD_DONE = True
                RetVal = CmdSuccess
            End If
        End While
        'D.FindDAQs(USB1208LS_Devices)
        'MsgBox(USB1208LS_Devices.Count.ToString)
    End Sub
End Class
