Imports System.IO.Ports
Imports System.Text.RegularExpressions

Public Class BK1856D
    Private SP As SerialPort
    Private _ErrorMsg As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Function Open(ByVal SerialPortName As String) As Boolean
        Dim readingResult As Boolean = False
        Dim reading As Double
        Dim startTime As DateTime

        'Cycle comport needed for Windows to detect the port properly
        Dim ft232 As New FT232R
        Dim status = ft232.CycleComport(SerialPortName)
        If Not status Then
            Form1.AppendText("Problem cycleComport BK1856D serial port")
            Return False
        End If

        System.Threading.Thread.Sleep(2000)

        SP = New SerialPort(SerialPortName, 9600, 0, 8, 1)
        SP.ReadTimeout = 1500
        SP.DtrEnable = True

        Try
            SP.Open()
        Catch ex As Exception
            _ErrorMsg = "SP.Open failed"
            _ErrorMsg += vbCr + ex.ToString
            Return False
        End Try
        'SP.NewLine = Chr(13)
        Try
            System.Threading.Thread.Sleep(1000)
            SP.Write("N0" + Chr(13))
            System.Threading.Thread.Sleep(1000)
            SP.Write("H0" + Chr(13))
            System.Threading.Thread.Sleep(1000)
            SP.Write("G2" + Chr(13))
        Catch ex As Exception
            _ErrorMsg = "Error writing to BK 1856D " + vbCr + ex.ToString
            Try
                SP.Close()
            Catch exx As Exception
                MessageBox.Show("Exception : " + ex.Message + " while trying to open BK1856D", "Exception in reading loop")
            End Try
            Return False
        End Try
        SP.NewLine = Chr(13)
        reading = 0
        startTime = Now
        While (reading = 0 And Now.Subtract(startTime).TotalSeconds < 20)
            Application.DoEvents()
            System.Threading.Thread.Sleep(2000)
            readingResult = Read(reading)
            If Not readingResult Then
                Try
                    _ErrorMsg = "Reading call returned failure. Closing Port"
                    SP.Close()
                Catch ex As Exception
                    MessageBox.Show("Exception : " + ex.Message + " while trying to open BK1856D", "Exception in reading loop")
                End Try
                Return False
            End If
        End While

        If reading = 0 Then
            _ErrorMsg = "Timeout waiting for reading BK 1856D to be non zero"
            SP.Close()
            Return False
        End If

        Return True
    End Function

    Function Read(ByRef reading As Double) As Boolean
        Dim readingStr As String
        Dim Fields() As String
        Dim Units As String
        Dim prev_reading As Double
        Dim retryCnt As Integer = 0

        Try
            SP.ReadExisting()
        Catch ex As Exception

        End Try

        prev_reading = 666.666
        reading = 0
        While (Not reading = prev_reading) And (retryCnt < 10)
            prev_reading = reading
            Try
                System.Threading.Thread.Sleep(2000)
                SP.Write("D1" + Chr(13))
                System.Threading.Thread.Sleep(2000)
                readingStr = SP.ReadLine()
            Catch ex As Exception
                _ErrorMsg = "Problem reading 1856D" + vbCr + ex.ToString
                Return False
            End Try

            readingStr = readingStr.Trim()
            _ErrorMsg = "FREQ Counter String returned : " + readingStr
            If readingStr = "0" Then
                reading = 0
                Units = "Hz"
            ElseIf (Regex.IsMatch(readingStr, "\d+\.\d+\s+\w+$")) Then
                Fields = Regex.Split(readingStr, "(\d+\.\d+)\s+(\w+)")
                reading = CDbl(Fields(1))
                Units = Fields(2)
            Else
                _ErrorMsg = "Can't decipher 1856D reading '" + readingStr + "'"
                Return False
            End If

            retryCnt += 1
        End While

        If Not reading = prev_reading Then
            _ErrorMsg = "1856D measurement unstable"
            Return False
        End If

        If Units = "Hz" Then
            reading = reading * 1.0
        ElseIf Units = "kHz" Then
            reading = reading * 1000
        ElseIf Units = "mHz" Then
            reading = reading * 1000000
        Else
            _ErrorMsg = "Can't decipher 1856D reading '" + readingStr + "'"
            Return False
        End If
        Return True
    End Function

    Function Close() As Boolean
        Try
            SP.Close()
        Catch ex As Exception
            _ErrorMsg = "Problem closing COM connection to 1856D" + vbCr + ex.ToString
            Return False
        End Try
        Return True
    End Function

End Class
