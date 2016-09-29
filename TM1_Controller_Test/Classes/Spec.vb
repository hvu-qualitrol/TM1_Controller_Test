Public Class Spec
    Private _Min As Double
    Private _Max As Double
    Private _Expected As String
    Private _Actual As String
    Private _ErrorMsg As String = ""
    Private _InSpec As Boolean

    Public Property Min() As Double
        Set(value As Double)
            _Min = value
        End Set
        Get
            Return _Min
        End Get
    End Property

    Public Property Max() As Double
        Set(value As Double)
            _Max = value
        End Set
        Get
            Return _Max
        End Get
    End Property

    Public Property Expected() As String
        Set(value As String)
            _Expected = value
        End Set
        Get
            Return _Expected
        End Get
    End Property

    Public Property Actual() As String
        Set(value As String)
            _Actual = value
        End Set
        Get
            Return _Actual
        End Get
    End Property

    Public Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Property InSpec() As Boolean
        Set(value As Boolean)
            _InSpec = value
        End Set
        Get
            Return _InSpec
        End Get
    End Property

    Public Sub New(ByVal Min As Double, ByVal Max As Double)
        _Min = Min
        _Max = Max
    End Sub

    Public Sub New(ByVal Expected As String)
        _Expected = Expected
    End Sub

    Public Function CompareToSpec(ByVal Value As Double)
        _Actual = Value.ToString
        If Value < _Min Or Value > _Max Then
            _ErrorMsg = " Expected between " + _Max.ToString + " and " + _Min.ToString
            _InSpec = False
            Return False
        End If
        _InSpec = True
        Return True
    End Function
End Class
