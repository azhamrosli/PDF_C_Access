Public Class MyListItem
    Private mstrText As String
    Private mstrValue As String
    Public Sub New()

    End Sub

    Public Sub New(ByVal strText As String, _
      ByVal strValue As String)
        mstrValue = strValue
        mstrText = strText
    End Sub

    Property Value() As String
        Get
            Return (mstrValue)
        End Get
        Set(ByVal Text As String)
            mstrValue = Text
        End Set
    End Property

    Property Text() As String
        Get
            Return (mstrText)
        End Get
        Set(ByVal Text As String)
            mstrText = Text
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return (mstrText)
    End Function
End Class
