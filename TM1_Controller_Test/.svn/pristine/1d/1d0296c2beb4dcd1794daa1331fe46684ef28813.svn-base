Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class DB
    Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader

    Private _ErrorMsg As String = ""

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function Connection() As SqlConnection
        'Return New SqlConnection("workstation id=PDX-001-703K;packet size=4096;user id=standard;data source=PDX-BAK-01;persist security info=True;initial catalog=Serveron Operation;password=password")        
        Dim SqlPacketSize As String
        Dim SqlUserId As String
        Dim SqlServerName As String
        Dim SqlPersistSecurity As String
        Dim SqlDataBaseName As String
        Dim SqlUserPassword As String

        SqlPacketSize = "packet size=" + Convert.ToString(configurationAppSettings.GetValue("SqlPacketSize", GetType(String)))
        SqlUserId = "user id=" + Convert.ToString(configurationAppSettings.GetValue("SqlUserId", GetType(String)))
        SqlServerName = "data source=" + Convert.ToString(configurationAppSettings.GetValue("SqlServerName", GetType(String)))
        SqlPersistSecurity = "persist security info=" + Convert.ToString(configurationAppSettings.GetValue("SqlPersistSecurity", GetType(String)))
        SqlDataBaseName = "initial catalog=" + Convert.ToString(configurationAppSettings.GetValue("SqlDataBaseName", GetType(String)))
        SqlUserPassword = "password=" + Convert.ToString(configurationAppSettings.GetValue("SqlUserPassword", GetType(String)))

        ConnectString = String.Format("{0};{1};{2};{3};{4};{5}", SqlPacketSize, SqlUserId, SqlServerName, SqlPersistSecurity, SqlDataBaseName, SqlUserPassword)
        Return New SqlConnection(ConnectString)
    End Function

    Public Function AddLinkData(ByVal NewLinkInfo As Hashtable) As Boolean
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim LinkInfo As Hashtable
        Dim UpdateNeeded As Boolean = False

        If Not GetLinkData(NewLinkInfo("TM1"), AssemblyType.TM1, LinkInfo) Then
            Return False
        End If
        If LinkInfo.Count > 0 Then
            For Each k In NewLinkInfo.Keys
                If Not LinkInfo(k) = NewLinkInfo(k) Then
                    UpdateNeeded = True
                End If
            Next
            If Not UpdateNeeded Then
                Return True
            End If
        End If
        If UpdateNeeded Then
            For Each k In LinkInfo.Keys
                If Not NewLinkInfo.Contains(k) And Not k = "Link_Time" Then
                    NewLinkInfo.Add(k, LinkInfo(k))
                End If
            Next
            If Not InvalidateLinkData(NewLinkInfo("TM1")) Then
                Return False
            End If
        End If
        If Not InsertLinkData(NewLinkInfo) Then
            Return False
        End If

        Return True
    End Function

    Public Function AddTm102LinkData(ByVal tm101Sn As String, ByVal tm102Sn As String) As Boolean
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim linkInfo As Hashtable = New Hashtable
        Dim UpdateNeeded As Boolean = False

        ' Fail if cannot get the link info on the specified TM101 unit
        If Not GetLinkData(tm101Sn, AssemblyType.TM1, linkInfo) Then
            MsgBox("Failed to get link info on " + tm101Sn)
            Return False
        End If
        If linkInfo.Count = 0 Then
            MsgBox("No link info on " + tm101Sn)
            Return False
        End If

        ' Change the TM101 to TM102 serial number
        linkInfo("TM1") = tm102Sn
        linkInfo.Remove("Link_Time")

        ' Insert the link info to the database
        If Not InsertLinkData(linkInfo) Then
            MsgBox("Failed to insert to database the link info of " + tm102Sn)
            Return False
        End If

        Return True
    End Function

    Public Function InsertLinkData(ByVal LinkInfo As Hashtable)
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim SqlCmd As SqlCommand
        Dim rowsAffected As Integer

        Try
            SqlCmd = New SqlCommand()
            With SqlCmd
                .CommandType = CommandType.Text
                .Connection = sqlConn
                .CommandText = "INSERT TM1_Link_Table (Link_Time"
                For Each k In LinkInfo.Keys
                    .CommandText += ", " + k
                Next
                .CommandText += ") values('" + Now.ToString + "'"
                For Each k In LinkInfo.Keys
                    .CommandText += ", '" + LinkInfo(k) + "'"
                Next
                .CommandText += ")"
            End With
            sqlConn.Open()
            rowsAffected = SqlCmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrorMsg = ex.ToString
            Return False
        Finally
            sqlConn.Close()
        End Try

        If Not rowsAffected > 0 Then
            _ErrorMsg = "Problem inserting data into TM1_Link_Table" + vbCr
            _ErrorMsg = "cmd:  " + SqlCmd.CommandText
            Return False
        End If

        Return True
    End Function

    Public Function InvalidateLinkData(ByVal SN As String)
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim rowsAffected As Integer

        Try
            SqlCmd.Connection = sqlConn
            SqlCmd.CommandText = "UPDATE TM1_Link_Table set valid='0' where TM1='" + SN + "'"
            sqlConn.Open()
            rowsAffected = SqlCmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrorMsg = ex.ToString
            Return False
        Finally
            sqlConn.Close()
        End Try

        If Not rowsAffected > 0 Then
            _ErrorMsg = "Problem updating link table"
            Return False
        End If

        Return True
    End Function

    Public Function GetLinkData(ByVal SN As String, ByVal AT As AssemblyType, ByRef LinkInfo As Hashtable) As Boolean
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim DtTemp As New DataTable

        LinkInfo = New Hashtable

        Try
            SqlCmd.Connection = sqlConn
            SqlCmd.CommandText = "Select TOP 1 * from TM1_Link_Table "
            Select Case AT
                Case AssemblyType.CONTROLLER_BOARD
                    SqlCmd.CommandText += "where CONTROLLER_BOARD='" + SN + "' "
                Case AssemblyType.H2SCAN
                    SqlCmd.CommandText += "where H2SCAN='" + SN + "' "
                Case AssemblyType.TM1
                    SqlCmd.CommandText += "where TM1='" + SN + "' "
                Case AssemblyType.DISPLAY_KIT
                    SqlCmd.CommandText += "where DISPLAY_KIT='" + SN + "' "
            End Select
            SqlCmd.CommandText += "AND valid = 1 order by Link_Time desc"
            sqlConn.Open()
            SqlAdapter.SelectCommand = SqlCmd
            If (SqlAdapter.Fill(DtTemp) = 1) Then
                LinkInfo.Add("TM1", DtTemp.Rows(0).Item("TM1"))
                If DtTemp.Rows(0).Item("H2SCAN") IsNot DBNull.Value Then
                    LinkInfo.Add("H2SCAN", DtTemp.Rows(0).Item("H2SCAN"))
                End If
                If DtTemp.Rows(0).Item("CONTROLLER_BOARD") IsNot DBNull.Value Then
                    LinkInfo.Add("CONTROLLER_BOARD", DtTemp.Rows(0).Item("CONTROLLER_BOARD"))
                End If
                If DtTemp.Rows(0).Item("PUMP") IsNot DBNull.Value Then
                    LinkInfo.Add("PUMP", DtTemp.Rows(0).Item("PUMP"))
                End If
                If DtTemp.Rows(0).Item("POWER_SUPPLY") IsNot DBNull.Value Then
                    LinkInfo.Add("POWER_SUPPLY", DtTemp.Rows(0).Item("POWER_SUPPLY"))
                End If
                If DtTemp.Rows(0).Item("DISPLAY_KIT") IsNot DBNull.Value Then
                    LinkInfo.Add("DISPLAY_KIT", DtTemp.Rows(0).Item("DISPLAY_KIT"))
                End If
                If DtTemp.Rows(0).Item("Link_Time") IsNot DBNull.Value Then
                    LinkInfo.Add("Link_Time", DtTemp.Rows(0).Item("Link_Time"))
                End If
            End If
        Catch ex As Exception
            _ErrorMsg = "Problem getting link data" + vbCr
            _ErrorMsg += "Cmd:  " + SqlCmd.CommandText + vbCr
            _ErrorMsg += ex.ToString
            Return False
        Finally
            sqlConn.Close()
        End Try

        Return True
    End Function

    Public Function AddH2scanData(ByVal H2SCAN_Data As Hashtable) As Boolean
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim DtTemp As New DataTable
        Dim DataExists As Boolean = False

        Try
            SqlCmd.Connection = sqlConn
            SqlCmd.CommandText = "Select TOP 1 * from TM1_H2SCAN_Data where Product_Serial='" + H2SCAN_Data("Sensor Product Serial")
            sqlConn.Open()
            SqlAdapter.SelectCommand = SqlCmd
            If (SqlAdapter.Fill(DtTemp) = 1) Then
                If DtTemp.Rows.Count > 0 Then
                    DataExists = True
                End If
            End If
        Catch ex As Exception
            _ErrorMsg = "Problem check H2SCAN data" + vbCr
            _ErrorMsg += ex.ToString
            sqlConn.Close()
            Return False
        End Try

        If DataExists Then
            SqlCmd.CommandText = "Update TM1_H2SCAN_Data "
        End If

        Return True
    End Function

    Public Function Pass(ByVal SN As String, ByVal AL As String, ByVal TestTime As Integer)
        Dim sqlConn As SqlConnection = Connection()
        Dim sqlCmd As SqlCommand
        Dim rowsAffected As Integer
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim TestID As Integer
        Dim DtTemp As New DataTable

        Try
            sqlCmd = New SqlCommand()
            With sqlCmd
                .CommandType = CommandType.Text
                .CommandText = "Insert TM1_Test_History (SerialNumber, PassFail, AssemblyLevel, TestTime, TestDate) " +
                    "values ('" + SN + "', 1, '" + AssemblyLevel + "', " + TestTime.ToString + ", '" + Now.ToString + "')"
                .Connection = sqlConn
            End With
            sqlConn.Open()
            rowsAffected = sqlCmd.ExecuteNonQuery()
            If Not rowsAffected = 1 Then
                _ErrorMsg = "Problem logging test history for " + SN + vbCr
                _ErrorMsg += "CMD:  " + sqlCmd.CommandText
                sqlConn.Close()
                Return False
            End If

            With sqlCmd
                .CommandType = CommandType.Text
                .CommandText = "SELECT @@IDENTITY AS 'Identity'"
                .Connection = sqlConn
            End With
            SqlAdapter.SelectCommand = sqlCmd
            SqlAdapter.Fill(DtTemp)
            TestID = DtTemp.Rows(0).Item("Identity")
        Catch ex As Exception
            _ErrorMsg = "Problem logging test history for " + SN + vbCr
            _ErrorMsg += "CMD:  " + sqlCmd.CommandText + vbCr
            _ErrorMsg += ex.ToString
            sqlConn.Close()
            Return False
        End Try
        sqlConn.Close()

        If Not LogMeasurements(TestID) Then
            Form1.AppendText("Problem logging measurement data to database")
            Return False
        End If

        Return True
    End Function

    Public Function Fail(ByVal SN As String, ByVal AL As String, ByVal TestTime As Integer, ByVal TestName As String, _
                         ByVal TestDetails As String)
        Dim sqlConn As SqlConnection = Connection()
        Dim sqlCmd As SqlCommand = Nothing
        Dim rowsAffected As Integer
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim TestID As Integer
        Dim DtTemp As New DataTable
        Dim FirstLine As Boolean = True

        Try
            sqlCmd = New SqlCommand()
            With sqlCmd
                .CommandType = CommandType.Text
                .CommandText = "Insert TM1_Test_History (SerialNumber, PassFail, AssemblyLevel, TestTime, TestDate) " +
                    "values ('" + SN + "', 0, '" + AssemblyLevel + "', " + TestTime.ToString + ", '" + Now.ToString + "')"
                .Connection = sqlConn
            End With
            sqlConn.Open()
            rowsAffected = sqlCmd.ExecuteNonQuery()
            If Not rowsAffected = 1 Then
                _ErrorMsg = "Problem logging test history for " + SN + vbCr
                _ErrorMsg += "CMD:  " + sqlCmd.CommandText
                sqlConn.Close()
                Return False
            End If

            With sqlCmd
                .CommandType = CommandType.Text
                .CommandText = "SELECT @@IDENTITY AS 'Identity'"
                .Connection = sqlConn
            End With
            SqlAdapter.SelectCommand = sqlCmd
            SqlAdapter.Fill(DtTemp)
            TestID = DtTemp.Rows(0).Item("Identity")

            TestDetails = Regex.Replace(TestDetails, "\'", "^")
            With sqlCmd
                .CommandType = CommandType.Text
                .CommandText = "Insert TM1_Failure_Data (TestID, TestName, TestDetails) Values (" + TestID.ToString + ", '" + TestName + "', "
                For Each Line In Regex.Split(TestDetails, System.Environment.NewLine)
                    If Not FirstLine Then
                        .CommandText += " + '" + Line + "' + CHAR(13) + CHAR(10)"
                    Else
                        .CommandText += "'" + Line + "' + CHAR(13) + CHAR(10)"
                        FirstLine = False
                    End If
                Next
                .CommandText += ")"
                .Connection = sqlConn
            End With
            rowsAffected = sqlCmd.ExecuteNonQuery()
            If Not rowsAffected = 1 Then
                _ErrorMsg = "Problem logging failures details for " + SN + vbCr
                _ErrorMsg += "CMD:  " + sqlCmd.CommandText
                sqlConn.Close()
                Return False
            End If

        Catch ex As Exception
            _ErrorMsg = "Problem logging test history for " + SN + vbCr
            _ErrorMsg += "CMD:  " + sqlCmd.CommandText + vbCr
            _ErrorMsg += ex.ToString
            sqlConn.Close()
            Return False
        End Try
        sqlConn.Close()

        If Not LogMeasurements(TestID) Then
            Form1.AppendText("Problem logging measurement data to database")
            Return False
        End If

        Return True

    End Function

    Public Function LogMeasurements(ByVal TestID As Integer) As Boolean
        Dim sqlConn As SqlConnection = Connection()
        Dim sqlCmd As SqlCommand
        Dim rowsAffected As Integer
        Dim success As Boolean = True
        Dim PassFail As Integer

        Try
            sqlCmd = New SqlCommand()
            sqlConn.Open()

            For Each k In Specs.Keys
                If Specs(k).Actual = Nothing Then Continue For
                With sqlCmd
                    .CommandType = CommandType.Text
                    If Specs(k).InSpec Then
                        PassFail = 1
                    Else
                        PassFail = 0
                    End If
                    .CommandText = "Insert TM1_Test_Data (TestID, Date_Time, Serial_Number, SpecName, Result, PassFail) " +
                        "Values (" + TestID.ToString + ", '" + Now.ToString + "', '" + Form1.Serial_Number + "', '" + k + "', '" + Specs(k).Actual + "', " + PassFail.ToString + ")"
                    .Connection = sqlConn
                End With
                rowsAffected = sqlCmd.ExecuteNonQuery()
            Next
        Catch e As Exception
            Form1.AppendText("CMD:  " + sqlCmd.CommandText)
            Form1.AppendText(e.Message.ToString, False)
            success = False
        Finally
            sqlConn.Close()
        End Try
        If success Then
            Form1.AppendText("database.LogMeasurement success", False)
        End If
        Return success
    End Function

    Public Function GetLastTestResult(ByVal SN As String, ByVal AssemblyLevel As String, ByRef Passed As Boolean) As Boolean
        Dim SqlAdapter As New SqlClient.SqlDataAdapter
        Dim sqlConn As SqlConnection = Connection()
        Dim SqlCmd As New SqlClient.SqlCommand
        Dim DtTemp As New DataTable
        Dim TestResultFound As Boolean = False

        Passed = False
        Try
            SqlCmd.Connection = sqlConn
            SqlCmd.CommandText = "Select TOP 1 * from TM1_Test_History where SerialNumber='" + SN + "' and AssemblyLevel='" + AssemblyLevel + "' order by TestDate desc"
            sqlConn.Open()
            SqlAdapter.SelectCommand = SqlCmd
            If (SqlAdapter.Fill(DtTemp) = 1) Then
                If DtTemp.Rows(0).Item("PassFail") = True Then
                    Passed = True
                End If
            End If
        Catch ex As Exception
            _ErrorMsg = "Problem getting test history for " + SN + vbCr
            _ErrorMsg += "Cmd:  " + SqlCmd.CommandText + vbCr
            _ErrorMsg += ex.ToString
            Return False
        Finally
            sqlConn.Close()
        End Try

        Return True
    End Function


    Public Function CheckPrevTests(ByVal SN As String, ByVal AssemblyLevel As String) As Boolean
        Dim ControllerBoardSN As String
        Dim LinkInfo As Hashtable
        Dim Passed As Boolean

        If AssemblyLevel = "SUB_ASSEMBLY" Or AssemblyLevel = "W223_REWORK" Then
            'If Not GetLinkData(SN, AssemblyType.TM1, LinkInfo) Then
            '    Return False
            'End If
            'If Not LinkInfo.Contains("CONTROLLER_BOARD") Then
            '    _ErrorMsg = "controller board SN not found linked to " + SN
            '    Return False
            'End If
            'ControllerBoardSN = LinkInfo("CONTROLLER_BOARD")
            ControllerBoardSN = Form1.Controller_Board_SN_Entry.Text
            If Not GetLastTestResult(ControllerBoardSN, "CONTROLLER_BOARD", Passed) Then
                Form1.AppendText(_ErrorMsg)
                Return False
            End If
            If Not Passed Then
                Form1.AppendText("Passing result for controller board test not found")
                Return False
            End If

            If Not GetLastTestResult(SN, "LEAK", Passed) Then
                Form1.AppendText(_ErrorMsg)
                Return False
            End If
            If Not Passed Then
                Form1.AppendText("LEAK test passing results not seen")
                Return False
            End If
        End If

        Return True
    End Function
End Class
'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           