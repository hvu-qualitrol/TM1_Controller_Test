Imports System.Text.RegularExpressions
Imports System.IO
Imports System.IO.File

Public Class CommonLib
    Shared _ErrorMsg As String = ""

    Shared Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Shared Sub Delay(ByVal delay_time As Int16, Optional ByVal DisplayTimeout As Boolean = False)
        Dim Counts As Integer = delay_time / 0.2
        Dim Count As Integer = 0

        While Count < Counts
            If DisplayTimeout Then
                Form1.TimeoutLabel.Text = Math.Round((Counts - Count) * 0.2).ToString
            End If
            System.Threading.Thread.Sleep(200)
            Count += 1
            Application.DoEvents()
        End While
    End Sub

    Shared Function ZuluToUTC(ByVal ZuluTime As String, ByRef UTC_DateTime As DateTime) As Boolean
        Dim TimeFields() As String

        Try
            TimeFields = Regex.Split(ZuluTime, "-|T|Z|:")
            UTC_DateTime = New DateTime(CInt(TimeFields(0)), CInt(TimeFields(1)), CInt(TimeFields(2)),
                                                    CInt(TimeFields(3)), CInt(TimeFields(4)), CInt(TimeFields(5)))
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Shared Function GetLocalTime(ByRef CurrentTime As String, Optional ByVal Fmt As String = "STD") As Boolean
        Dim ServerTime As DateTime = Now
        Dim ServerTimeStr As String

        If Fmt = "ZULU" Then
            ServerTimeStr = ServerTime.ToString
            CurrentTime = Format(ServerTime.ToUniversalTime, "yyyy-MM-dd\THH:mm:ss\Z")
        ElseIf Fmt = "UTC" Then
            CurrentTime = ServerTime.ToUniversalTime.ToString
        Else
            CurrentTime = ServerTime.ToString
        End If

        Return True
    End Function

    Shared Function GetNetworkTime(ByRef CurrentTime As String, Optional ByVal Fmt As String = "STD") As Boolean
        Dim p As New Process
        Dim Line As String
        Dim startTime As String = Now
        Dim ServerTimeStr As DateTime
        Dim GotTime As Boolean = False
        Dim ServerTime As DateTime

        p.StartInfo.UseShellExecute = False
        p.StartInfo.CreateNoWindow = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.FileName = "net"
        p.StartInfo.Arguments = "time \\d600loaner"

        p.Start()
        startTime = Now
        While (Not p.HasExited) And (Now.Subtract(startTime).TotalSeconds < 30)
            Application.DoEvents()
        End While
        If Not p.HasExited Then
            CurrentTime = "Timeout getting network time"
            p.Close()
            Return False
        End If
        Line = p.StandardOutput.ReadLine

        While (Not Line = Nothing)
            Console.WriteLine(Line)
            If (Line.Contains("Current time")) Then
                ServerTimeStr = Line.Substring(Line.IndexOf(" is") + 4)
                ServerTime = ServerTimeStr
                ServerTime.AddSeconds(Now.Subtract(startTime).TotalSeconds)
                ServerTimeStr = ServerTime.ToString
                'MsgBox(ServerTime.ToString + vbCr + Now.ToString + vbCr + Now.Subtract(startTime).TotalSeconds.ToString)
                If Fmt = "ZULU" Then
                    CurrentTime = Format(ServerTimeStr, "yyyy-MM-dd\THH:mm:ss\Z")
                ElseIf Fmt = "UTC" Then
                    CurrentTime = ServerTimeStr.ToUniversalTime.ToString()
                Else
                    CurrentTime = ServerTimeStr.ToString
                End If
                GotTime = True
            End If
            Line = p.StandardOutput.ReadLine
        End While
        If Not GotTime Then
            CurrentTime = "Problem getting server time"
        End If

        p.Close()

        Return GotTime
    End Function

    Shared Function ExportDataTableToCSV(ByVal DT As DataTable, ByVal Filename As String) As Boolean
        Dim CSVFile As FileStream
        Dim CSVFileWriter As StreamWriter
        Dim column As DataColumn
        Dim FirstCol As Boolean
        Dim row As DataRow

        CSVFile = New FileStream(Filename, FileMode.Create, FileAccess.Write)
        CSVFileWriter = New StreamWriter(CSVFile)

        FirstCol = True
        For Each column In DT.Columns
            If Not FirstCol Then
                CSVFileWriter.Write(",")
            Else
                FirstCol = False
            End If
            CSVFileWriter.Write(column.ColumnName)
        Next
        CSVFileWriter.Write(vbCr)

        For Each row In DT.Rows
            For i = 0 To DT.Columns.Count - 1
                CSVFileWriter.Write(row.Item(i).ToString)
                If i < DT.Columns.Count - 1 Then
                    CSVFileWriter.Write(",")
                End If
            Next
            CSVFileWriter.Write(vbCr)
        Next

        CSVFileWriter.Close()
        CSVFile.Close()


        Return True
    End Function

    'Shared Function FindThumbDrive(ByRef DriveLetter) As Boolean
    '    Dim allDrives() As DriveInfo
    '    Dim d As DriveInfo
    '    Dim ThumbDriveCnt As Integer = 0
    '    Dim startTime As DateTime = Now

    '    allDrives = DriveInfo.GetDrives
    '    For Each d In allDrives
    '        If (d.DriveType = DriveType.Removable) Then
    '            ThumbDriveCnt += 1
    '            DriveLetter = d.Name
    '            While Not d.IsReady And Now.Subtract(startTime).TotalSeconds < 20

    '            End While
    '            If Not d.IsReady Then
    '                _ErrorMsg = "Drive is not ready"
    '                Return False
    '            End If
    '        End If
    '    Next

    '    If Not ThumbDriveCnt = 1 Then
    '        _ErrorMsg = "Found " + ThumbDriveCnt.ToString + " thumb drives, expected 1"
    '        Return False
    '    End If

    '    Return True
    'End Function
End Class
