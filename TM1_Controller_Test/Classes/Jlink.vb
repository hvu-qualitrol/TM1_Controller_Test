Imports System.IO
Imports System.IO.File
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class Jlink
    Private JlinkOutputFile As FileStream
    Private JlinkOutputFileWriter As StreamWriter
    Private JlinkStderrFile As FileStream
    Private JlinkStderrFileWriter As StreamWriter
    Private VTarget_min As Double = 3.2
    Private VTarget_max As Double = 3.4
    Private _results As String = ""

    Property Results() As String
        Get
            Return _results
        End Get
        Set(value As String)

        End Set
    End Property

    Private Function FindJlink(ByRef JlinkPath) As Boolean
        Dim PathHash As New Hashtable
        Dim version As String
        Dim versions As New List(Of String)
        Dim path As String
        Dim SEGGER_DIR As String

        Form1.AppendText("FindJlink() start...")

        If Directory.Exists("C:\Program Files\SEGGER\") Then
            SEGGER_DIR = "C:\Program Files\SEGGER\"
        ElseIf Directory.Exists("C:\Program Files (x86)\SEGGER\") Then
            SEGGER_DIR = "C:\Program Files (x86)\SEGGER\"
        Else
            _results = "Can't find SEGGER directory"
            Return False
        End If

        Form1.AppendText("Directory.GetFiles() JLink.exe from " + SEGGER_DIR)
        For Each path In Directory.GetFiles(SEGGER_DIR, "JLink.exe", SearchOption.AllDirectories)
            Try
                version = Regex.Split(path, "JLinkARM_V(\w+)\\")(1)
                versions.Add(version)
                PathHash.Add(version, path)
            Catch e As Exception
                Form1.AppendText("Caught " + e.Message)
            End Try
        Next

        If versions.Count < 1 Then
            _results = "Cant find JLink.exe, SEGGER SW may not be installed"
            Return False
        End If
        versions.Sort()
        version = versions(versions.Count - 1)
        JlinkPath = PathHash(version)

        Form1.AppendText("FindJlink() complete!")

        Return True
    End Function

    Private Function RunJlink(ByVal UniqueExpectedStrings() As String, ByVal jlink_script As String, ByVal Timeout As Integer)
        Dim p As New Process
        Dim startTime As DateTime
        Dim JlinkResultFile As FileStream
        Dim JlinkResultFileReader As StreamReader
        Dim ExpectedStrings(10) As String
        Dim ExpectedIndex As Integer
        Dim ExpectedCnt As Integer
        Dim Line As String
        Dim VTarget As Double
        Dim reportTime As DateTime
        Dim JlinkPath As String

        Form1.AppendText("RunJLink() start...")

        ExpectedStrings(0) = "Script file read successfully"
        ExpectedStrings(1) = "Cortex-M4 identified"
        UniqueExpectedStrings.CopyTo(ExpectedStrings, 2)
        ExpectedCnt = 2 + UniqueExpectedStrings.Length

        Form1.AppendText("RunJLink().FindJLink() at " + JlinkPath)
        If Not FindJlink(JlinkPath) Then
            Return False
        End If

        Form1.AppendText("Create process based on " + JlinkPath)
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.CreateNoWindow = True
        ' p.StartInfo.FileName = "C:\Program Files\SEGGER\JLinkARM_V436e\JLink.exe"
        ' p.StartInfo.FileName = "C:\Program Files\SEGGER\JLinkARM_V450k\JLink.exe"
        p.StartInfo.FileName = JlinkPath
        p.StartInfo.Arguments = jlink_script

        JlinkOutputFile = New FileStream("\temp\Jlink_output.txt", FileMode.Create, FileAccess.Write)
        JlinkOutputFileWriter = New StreamWriter(JlinkOutputFile)
        JlinkStderrFile = New FileStream("\temp\Jlink_stderr.txt", FileMode.Create, FileAccess.Write)
        JlinkStderrFileWriter = New StreamWriter(JlinkStderrFile)

        AddHandler p.OutputDataReceived, AddressOf JlinkOutputHandler
        AddHandler p.ErrorDataReceived, AddressOf JlinkStderrHandler

        Form1.AppendText("Start the created process...")
        Try
            p.Start()
        Catch ex As Exception
            _results += "p.start " + ex.ToString
            Return False
        End Try
        p.BeginOutputReadLine()
        p.BeginErrorReadLine()

        Form1.AppendText("Wait for the process complete...")

        startTime = Now
        reportTime = Now
        While (Not p.HasExited And (Now.Subtract(startTime).TotalSeconds) < Timeout)
            Thread.Sleep(500)
            Application.DoEvents()
            If (Now.Subtract(reportTime).TotalSeconds > 4) Then
                reportTime = Now
                Form1.AppendText("Timeout in " + (Timeout - Now.Subtract(startTime).TotalSeconds).ToString, True)
            End If
        End While

        If Not p.HasExited Then
            p.Kill()
            Form1.AppendText("Timeout", True)
            Return False
        End If
        Form1.AppendText("The process complete!")

        JlinkOutputFileWriter.Close()
        JlinkOutputFile.Close()
        JlinkStderrFileWriter.Close()
        JlinkStderrFile.Close()
        JlinkResultFile = New FileStream("\temp\Jlink_output.txt", FileMode.Open, FileAccess.Read)
        JlinkResultFileReader = New StreamReader(JlinkResultFile)

        Form1.AppendText("Parse the results...")

        ExpectedIndex = 0
        Line = JlinkResultFileReader.ReadLine()
        While (Not Line Is Nothing)
            If Not ExpectedStrings(ExpectedIndex) = Nothing Then
                If Line.Contains(ExpectedStrings(ExpectedIndex)) Then
                    ExpectedIndex += 1
                End If
            End If
            Form1.AppendText(Line, True)
            If Regex.IsMatch(Line, "^VTarget\s+=\s+\d+\.\d+") Then
                VTarget = Regex.Split(Line, "(\d+\.\d+)")(1)
            End If

            Line = JlinkResultFileReader.ReadLine()
        End While

        JlinkResultFileReader.Close()
        JlinkResultFile.Close()

        Form1.AppendText("ExpectedIndex = " + ExpectedIndex.ToString, True)
        If Not ExpectedIndex = ExpectedCnt Then
            Form1.AppendText("Expected :  " + ExpectedStrings(ExpectedIndex), True)
            Return False
        End If

        If VTarget < VTarget_min Or VTarget > VTarget_max Then
            Form1.AppendText("Expected VTarget between " + VTarget_min.ToString + " and " + VTarget_max.ToString, True)
            Return False
        End If

        Form1.AppendText("RunJLink() complete!")

        Return True
    End Function

    Public Function Verify(ByVal Version As String) As Boolean
        Dim md5sum_in As String
        Dim md5sum_out As String
        Dim fw_file As String
        Dim finfo As FileInfo
        Dim length As Long

        Form1.AppendText("Verify() start....")

        _results = ""
        Form1.AppendText("Verify(): Check file existing....")
        fw_file = BaseFWDirectory + "TM1" + "." + Version + ".bin"
        If Not Exists(fw_file) Then
            _results += fw_file + " does not exist" + vbCr
            Return False
        End If
        Form1.AppendText("Verify().SaveBin() start....")
        finfo = New FileInfo(fw_file)
        length = finfo.Length
        _results = "length = " + length.ToString + vbCr
        If Not SaveBin(length) Then
            _results += "Problem reading flash and saving to file"
            Return False
        End If

        Form1.AppendText("Compute md5sum_in checksum...")
        md5sum_in = MD5CalcFile(fw_file)
        _results += "checksum expected = " + md5sum_in + vbCr

        Form1.AppendText("Compute md5sum_out checksum...")
        md5sum_out = MD5CalcFile("C:\temp\out.bin")
        _results += "checksum actual = " + md5sum_out + vbCr

        Form1.AppendText("Compare checksums...")
        If Not md5sum_in = md5sum_out Then
            _results += "Checksum read back does not match checksum writte" + vbCr
            Return False
        End If
        Form1.AppendText("Verify() complete")


        Return True
    End Function

    Public Function Flash(ByVal Version As String) As Boolean
        Dim p As New Process
        Dim autoit_process As New Process
        Dim Timeout As Integer = 60
        Dim VTarget As Double = 0.0
        Dim JlinkFile As FileStream
        Dim JlinkFileWriter As StreamWriter
        Dim fw_file As String

        fw_file = BaseFWDirectory + "TM1" + "." + Version + ".bin"
        If Not Exists(fw_file) Then
            MsgBox(BaseFWDirectory + " does not exist")
            Return False
        End If

        JlinkFile = New FileStream(BaseFWDirectory + "TM1_flash.jlink", FileMode.Create, FileAccess.Write)
        JlinkFileWriter = New StreamWriter(JlinkFile)
        JlinkFileWriter.WriteLine("log \Temp\jlink_log.txt")
        JlinkFileWriter.WriteLine("Unlock Kinetis")
        JlinkFileWriter.WriteLine("speed 20000")
        JlinkFileWriter.WriteLine("exec device = MK60N512VLQ100")
        JlinkFileWriter.WriteLine("r")
        JlinkFileWriter.WriteLine("h")
        JlinkFileWriter.WriteLine("loadbin " + fw_file + ",0")
        JlinkFileWriter.WriteLine("qc")
        JlinkFileWriter.Close()
        JlinkFile.Close()

        'start autoit process that will close the popup dialog secure warning box from jlink if seen
        autoit_process.StartInfo.UseShellExecute = False
        autoit_process.StartInfo.RedirectStandardOutput = False
        autoit_process.StartInfo.RedirectStandardError = False
        autoit_process.StartInfo.CreateNoWindow = True
        ' autoit_process.StartInfo.FileName = "C:\autoit_scripts\coretex_secure"
        autoit_process.StartInfo.FileName = UtilitiesDirectory + "coretex_secure.exe"
        autoit_process.Start()
        System.Threading.Thread.Sleep(2000)

        Dim ExpectedStrings() As String = {"Flash programming performed",
                                           "Script processing completed"}
        ' If Not RunJlink(ExpectedStrings, "C:\NPI\TM1\firmware\TM1.jlink", 60) Then
        If Not RunJlink(ExpectedStrings, BaseFWDirectory + "TM1_flash.jlink", 60) Then
            Return False
        End If

        If Not autoit_process.HasExited Then
            autoit_process.Kill()
        End If

        Return True
    End Function

    Function SaveBin(ByVal length As Double)
        Form1.AppendText("Savebin() start...")

        Dim JlinkFile As FileStream
        Dim JlinkFileWriter As StreamWriter
        Dim ExpectedStrings() As String = {"Script processing completed"}

        Form1.AppendText("Write info into C:\Temp\out.bin...")

        JlinkFile = New FileStream("\Temp\TM1_savebin.jlink", FileMode.Create, FileAccess.Write)
        JlinkFileWriter = New StreamWriter(JlinkFile)
        JlinkFileWriter.WriteLine("speed 20000")
        JlinkFileWriter.WriteLine("exec device = MK60N512VLQ100")
        JlinkFileWriter.WriteLine("r")
        JlinkFileWriter.WriteLine("h")
        JlinkFileWriter.WriteLine("savebin C:\Temp\out.bin,0,0x" + Hex(length))
        JlinkFileWriter.WriteLine("qc")
        JlinkFileWriter.Close()
        JlinkFile.Close()

        Form1.AppendText("RunJLink() into C:\Temp\TM1_savebin.jlink...")

        'If Not RunJlink(ExpectedStrings, "C:\NPI\TM1\firmware\TM1_verify.jlink", 60) Then
        If Not RunJlink(ExpectedStrings, "C:\Temp\TM1_savebin.jlink", 60) Then
            Return False
        End If

        Form1.AppendText("Savebin() complete")

        Return True
    End Function

    Private Sub JlinkOutputHandler(sendingProcess As Object, outLine As DataReceivedEventArgs)
        JlinkOutputFileWriter.Write(Environment.NewLine + outLine.Data)
    End Sub

    Private Sub JlinkStderrHandler(sendingProcess As Object, outLine As DataReceivedEventArgs)
        JlinkStderrFileWriter.Write(Environment.NewLine + outLine.Data)
    End Sub

    Public Function MD5CalcFile(ByVal filepath As String) As String

        ' open file (as read-only)
        Using reader As New System.IO.FileStream(filepath, IO.FileMode.Open, IO.FileAccess.Read)
            Using md5 As New System.Security.Cryptography.MD5CryptoServiceProvider

                ' hash contents of this stream
                Dim hash() As Byte = md5.ComputeHash(reader)

                ' return formatted hash
                Return ByteArrayToString(hash)

            End Using
        End Using

    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String

        Dim sb As New System.Text.StringBuilder(arrInput.Length * 2)

        For i As Integer = 0 To arrInput.Length - 1
            sb.Append(arrInput(i).ToString("X2"))
        Next

        Return sb.ToString().ToLower

    End Function

    'Private Function Srec2Bin() As Boolean

    'End Function
End Class
