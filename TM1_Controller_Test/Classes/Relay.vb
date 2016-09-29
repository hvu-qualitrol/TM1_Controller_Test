Public Class Relay
    Private _COM As Integer
    Private _NO As Integer
    Private _NC As Integer

    Public Property COM As Integer
        Get
            Return _COM
        End Get
        Set(ByVal value As Integer)
            _COM = value
        End Set
    End Property

    Public Property NO As Integer
        Get
            Return _NO
        End Get
        Set(ByVal value As Integer)
            _NO = value
        End Set
    End Property

    Public Property NC As Integer
        Get
            Return _NC
        End Get
        Set(ByVal value As Integer)
            _NC = value
        End Set
    End Property

End Class
