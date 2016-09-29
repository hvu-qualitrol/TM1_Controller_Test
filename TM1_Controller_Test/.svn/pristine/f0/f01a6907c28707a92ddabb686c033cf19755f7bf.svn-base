Imports Microsoft.Office.Interop
Imports System.IO
Imports FTD2XX_NET

Partial Class Tests

    Public Shared Function ConvertFtdiSerialNumber() As Boolean
        Dim SF As New SerialFunctions

        tm101Sn = Form1.Serial_Number_Entry.Text
        tm102Sn = Form1.Controller_Board_SN_Entry.Text

        Dim ft232r As New FT232R
        Dim ftdiDevice As New FTDI

        Try
            If ftdiDevice.OpenBySerialNumber(tm102Sn) = FTDI.FT_STATUS.FT_OK Then
                Form1.AppendText("FTDI SN already set, skipping FTDI SN programming")
                ftdiDevice.Close()
                Return True
            End If

            If Not ftdiDevice.OpenBySerialNumber(tm101Sn) = FTDI.FT_STATUS.FT_OK Then
                Form1.AppendText("Problem opening FTDI device with " + tm101Sn, True)
                Return False
            End If

            If Not ft232r.SetSN(ftdiDevice, tm102Sn) Then
                Form1.AppendText("Problem setting ftdi device SN to " + tm102Sn)
                Form1.AppendText(ft232r.failure_message, True)
                Return False
            End If
            Form1.AppendText("SN set", True)

            If Not ftdiDevice.CyclePort() = FTDI.FT_STATUS.FT_OK Then
                Form1.AppendText("Problem cycling USB port", True)
                Return False
            End If
            CommonLib.Delay(30)

            ' Wait for the port to complete the cycle (reset) process
            Dim startTime As DateTime = Now
            Dim cycled As Boolean = False
            While Not cycled And Now.Subtract(startTime).TotalSeconds < 30
                CommonLib.Delay(3)
                If ftdiDevice.OpenBySerialNumber(tm102Sn) = FTDI.FT_STATUS.FT_OK Then
                    cycled = True
                End If
            End While
            ftdiDevice.Close()
            SerialPort = Nothing

        Catch ex As Exception
            If ftdiDevice.IsOpen Then ftdiDevice.Close()
            Form1.AppendText("ConvertFtdiSerialNumber() caught " + ex.ToString)
        End Try

        Return True

    End Function

    Public Shared Function ConvertConfigSerialNumber() As Boolean
        Dim pass As Boolean = True
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim Response As String = ""
        Dim Command As String
        Dim T As New TM1
        Dim VersionInfo As New Hashtable
        Dim DB As New DB

        tm101Sn = Form1.Serial_Number_Entry.Text
        tm102Sn = Form1.Controller_Board_SN_Entry.Text

        Try
            results = SF.Connect(SerialPort)
            If Not results.PassFail Then
                Return False
            End If

            If Not T.GetVersionInfo(VersionInfo) Then
                Return False
            End If
            If Not (VersionInfo.Contains("serial number")) Then
                Return False
            End If
            If VersionInfo("serial number") = tm102Sn Then
                Form1.AppendText("serial number already set")
                Return True
            End If

            Command = "config -S set SERIAL_NUMBER$ " + tm102Sn
            If Not SF.Cmd(SerialPort, Response, "config -S set SERIAL_NUMBER$ " + tm102Sn, 10) Then
                Form1.AppendText("Problem sending cmd '" + Command + "'")
                Form1.AppendText(SF.ErrorMsg)
                Return False
            End If

            System.Threading.Thread.Sleep(150)
            results = SF.Reboot(SerialPort)
            If Not results.PassFail Then
                Form1.AppendText("Problem rebooting")
                Form1.AppendText(SF.ErrorMsg)
                Return False
            End If

            If Not T.GetVersionInfo(VersionInfo) Then
                Return False
            End If
            If Not (VersionInfo.Contains("serial number")) Then
                Return False
            End If
            If Not VersionInfo("serial number") = tm102Sn Then
                Form1.AppendText("Expected 'serial number' = " + tm102Sn)
                Return False
            End If

        Catch ex As Exception
            Form1.AppendText("ConvertConfigSerialNumber() caught " + ex.ToString)
        End Try

        Return True
    End Function

    Public Shared Function ConvertTm101ToTm102() As Boolean
        Dim DB As New DB

        tm101Sn = Form1.Serial_Number_Entry.Text
        tm102Sn = Form1.Controller_Board_SN_Entry.Text


        Try
            ' Handle special functions for converting TM101 to TM102
            If AssemblyLevel = "TM101_TM102" And tm101Sn <> "" And tm102Sn <> "" Then
                ' Add link record to the database table
                If Not DB.AddTm102LinkData(tm101Sn, tm102Sn) Then Return False

                ' Convert TM101 to TM102 final report
                If Not ConvertWordReport(tm101Sn, tm102Sn) Then Return False

                ' If converted successfully, print the doc
                PrintTestReports(tm102Sn)
            End If
        Catch ex As Exception
            MsgBox("ConvertTm101ToTm102() caught " + ex.ToString)
            Return False
        End Try

        Return True

    End Function

    Private Shared Function ConvertWordReport(ByVal tm101Sn As String, ByVal tm102Sn As String)
        Dim pass As Boolean = True
        Dim tm101FilePath As String = "M:\Operations\Production\Operations Data\TM1\FINAL_TEST\" + tm101Sn + "\" + tm101Sn + " Test Summary Report.doc"
        Dim tm102FilePath As String = "M:\Operations\Production\Operations Data\TM1\TM101_TM102\" + tm102Sn + "\" + tm102Sn + " Test Summary Report.doc"
        Dim objWordApp As New Word.Application
        objWordApp.Visible = False

        ' Get the most current test summary report based on its timestamp
        Dim dir = New DirectoryInfo("M:\Operations\Production\Operations Data\TM1\FINAL_TEST\" + tm101Sn)
        Dim comparer As IComparer = New DateComparer()
        Dim files = dir.GetFileSystemInfos("TM101*.doc")
        Array.Sort(files, comparer)
        If files.Length > 0 Then
            tm101FilePath = files(0).FullName
        End If

        ' Bail out if the file does not exist
        If Not System.IO.File.Exists(tm101FilePath) Then
            MsgBox(tm101FilePath + " does not exist!")
            Return False
        End If

        Try
            If System.IO.File.Exists(tm102FilePath) Then System.IO.File.Delete(tm102FilePath)

            'Open an existing document.  
            Dim objDoc As Word.Document = objWordApp.Documents.Open(tm101FilePath)
            objDoc = objWordApp.ActiveDocument

            'Find and replace some text  
            'Replace 'VB' with 'Visual Basic'  
            'objDoc.Content.Find.Execute(FindText:="VB", ReplaceWith:="Visual Basic Express", Replace:=Word.WdReplace.wdReplaceAll)
            objDoc.Content.Find.Execute(FindText:=tm101Sn, ReplaceWith:=tm102Sn, Replace:=Word.WdReplace.wdReplaceAll)
            While objDoc.Content.Find.Execute(FindText:="  ", Wrap:=Word.WdFindWrap.wdFindContinue)
                objDoc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
            End While

            'Save and close the document  
            objDoc.SaveAs(tm102FilePath)
            'objDoc.Save()
            objDoc.Close()
            objDoc = Nothing
            objWordApp.Quit()
            objWordApp = Nothing
        Catch ex As Exception
            MsgBox("ConvertWordReport() caught " + ex.ToString)
            pass = False
        End Try

        Return pass

    End Function

    Private Shared Sub PrintTestReports(ByVal tm102Sn As String)
        Dim PD As New PrintDialog
        Dim app As Word.Application
        Dim m As Object = System.Reflection.Missing.Value
        Dim doc As Word.Document
        Dim filePath As String = "M:\Operations\Production\Operations Data\TM1\TM101_TM102\" + tm102Sn + "\" + tm102Sn + " Test Summary Report.doc"

        If Not System.IO.File.Exists(filePath) Then
            MsgBox("Test report " + filePath + " does not exist!")
            Exit Sub
        End If

        If PD.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        app = New Word.Application
        app.Visible = False
        Try
            app.WordBasic.FilePrintSetup(Printer:=PD.PrinterSettings.PrinterName, DoNotSetAsSysDefault:=1)
            doc = app.Documents.Open(filePath, m, m, m, m, m, m, m, m, m, m, m)
            app.PrintOut()
            app.Documents.Close()
            app.Quit()
            app = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If app IsNot Nothing Then
                app.Documents.Close()
                app.Quit()
            End If

        End Try
    End Sub

    Private Class DateComparer
        Implements System.Collections.IComparer

        Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim FileInfo1 As System.IO.FileInfo = DirectCast(info1, System.IO.FileInfo)
            Dim FileInfo2 As System.IO.FileInfo = DirectCast(info2, System.IO.FileInfo)

            Dim Date1 As DateTime = FileInfo1.CreationTime
            Dim Date2 As DateTime = FileInfo2.CreationTime

            If Date1 > Date2 Then Return -1
            If Date1 < Date2 Then Return 1

            Return 0
        End Function
    End Class

End Class
