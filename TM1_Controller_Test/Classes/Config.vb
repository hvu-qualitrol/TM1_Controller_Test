Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class config
    Private _Error_Message As String

    Property ErrorMsg() As String
        Get
            Return _Error_Message
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function ReadExpectedConfig(ByVal ConfigName As String, ByRef Config As Hashtable) As Boolean
        Dim res() As String
        Dim swriter As Stream
        Dim ConfigData As String
        Dim ConfigFile As String = "NOT FOUND"
        Dim Line As String
        Dim Fields() As String
        Dim Name, Value As String

        Config = New Hashtable

        res = Me.GetType.Assembly.GetManifestResourceNames()
        For i = 0 To UBound(res)
            If res(i).EndsWith(".txt") Then
                If res(i).Contains(ConfigName) Then
                    ConfigFile = res(i)
                End If
            End If
        Next
        If ConfigFile = "NOT FOUND" Then
            _Error_Message = "Can't find embedded resource file " + ConfigName
            Return False
        End If

        Try
            swriter = Me.GetType.Assembly.GetManifestResourceStream(ConfigFile)
            Dim Bytes(swriter.Length) As Byte
            swriter.Read(Bytes, 0, swriter.Length)
            ConfigData = Encoding.ASCII.GetString(Bytes)
            'ConfigData = Regex.Replace(ConfigData, "\?", "")
            ConfigData = Regex.Replace(ConfigData, "\?\?\?", "")
        Catch ex As Exception
            _Error_Message = "Problem reading embedded resource file " + ConfigName + vbCr
            _Error_Message += ex.ToString
            Return False
        End Try

        For Each Line In Regex.Split(ConfigData.Substring(0, ConfigData.Length - 1), "\r\n")
            Fields = Regex.Split(Line, "=")
            Try
                Name = Fields(0)
                Value = Fields(1)
            Catch ex As Exception
                _Error_Message = "Problem extracting name/value from " + Line
                Return False
            End Try
            Name = Regex.Replace(Name, "^\s+", "")
            Name = Regex.Replace(Name, "\s+$", "")
            Value = Regex.Replace(Value, "^\s+", "")
            Value = Regex.Replace(Value, "\s+$", "")
            Config.Add(Name, Value)
        Next

        Return True
    End Function
End Class
