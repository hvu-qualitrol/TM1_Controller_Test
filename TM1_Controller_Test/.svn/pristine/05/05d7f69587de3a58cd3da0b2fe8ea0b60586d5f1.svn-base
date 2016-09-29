Imports System.IO
Imports System.IO.MemoryMappedFiles

Public Class SharedMemoryMessaging
    Private _MAX_MESSAGE_LENGTH As Integer = 5000

    'Sub New()
    '    FtdiMM_File = MemoryMappedFile.CreateOrOpen(FtdiMM_FileName, _MAX_MESSAGE_LENGTH)
    'End Sub

    Public Function SendMessage(ByVal Message As String) As Boolean
        FtdiMM_File = MemoryMappedFile.CreateOrOpen(FtdiMM_FileName, _MAX_MESSAGE_LENGTH)
        If Message.Length > _MAX_MESSAGE_LENGTH Then
            Return False
        End If

        Using Writer = FtdiMM_File.CreateViewAccessor(0, _MAX_MESSAGE_LENGTH)
            Writer.WriteArray(Of Char)(0, Message, 0, Message.Length)
        End Using

        Return True
    End Function

    Public Function ReadMessage(ByRef Message As String) As Boolean
        Try
            Using File = MemoryMappedFile.OpenExisting(FtdiMM_FileName)
                Using reader = File.CreateViewAccessor(0, _MAX_MESSAGE_LENGTH)
                    Dim chars(_MAX_MESSAGE_LENGTH) As Char
                    reader.ReadArray(Of Char)(0, chars, 0, _MAX_MESSAGE_LENGTH)
                    Message = CStr(chars)
                End Using
            End Using
        Catch noFile As FileNotFoundException
            Return False
        Catch ex As Exception
            Return False
            Form1.AppendText(ex.ToString)
        End Try
        Return True
    End Function
End Class
