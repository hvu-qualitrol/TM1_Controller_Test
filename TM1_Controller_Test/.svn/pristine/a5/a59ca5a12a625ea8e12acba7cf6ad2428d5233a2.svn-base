Public Delegate Function MyDelFun() As Boolean

Public Class Test_Item
    Private _TestButton As Button
    Private _Function As MyDelFun
    Private _EntriesNeeded As New Hashtable
    Dim Entries() As String = {"Serial Number", "TMCOM1 COM Port", "Assembly Level"}

    Public Property Button() As Button
        Set(value As Button)
            _TestButton = value
        End Set
        Get
            Return _TestButton
        End Get
    End Property

    Public Property Handler() As MyDelFun
        Set(value As MyDelFun)
            _Function = value
        End Set
        Get
            Return _Function
        End Get
    End Property

    Public Property EntriesNeeded() As Hashtable
        Set(value As Hashtable)
            Dim Entry As String
            _EntriesNeeded = value
            For Each Entry In Entries
                If Not _EntriesNeeded.Contains(Entry) Then
                    _EntriesNeeded.Add(Entry, False)
                End If
            Next
        End Set
        Get
            Return _EntriesNeeded
        End Get
    End Property
End Class
