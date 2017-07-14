
Public Class BorangSelector
    Public Enum BorangEnum
        BorangC2008
        BorangC2009
        '===NgKL C2010.1==='
        BorangC2010
        '===NgKL C2010.1 End==='
        'weihong
        BorangC2011
        'endweihong
        'LEESH 04-FEB-2012
        BorangC2012
        'LEESH END
        BorangC2013
        'simkh 2014
        BorangC2014
        'simkh 2015 su8.1
        BorangC2015
        'azham
        BorangC2016
        'azham 2017
        BorangC2017
    End Enum

    Private Shared _Borang As BorangEnum = BorangEnum.BorangC2008
    Public Shared Property Borang() As BorangEnum
        Get
            Return _Borang
        End Get
        Set(ByVal value As BorangEnum)
            _Borang = value
        End Set
    End Property

    Private Shared _Year As String = "2008"
    Public Shared Property Year() As String
        Get
            Return _Year
        End Get
        Set(ByVal value As String)
            _Year = value
        End Set
    End Property

    Private Shared _RefNo As String = "1062026609"
    Public Shared Property RefNo() As String
        Get
            Return _RefNo
        End Get
        Set(ByVal value As String)
            _RefNo = value
        End Set
    End Property
End Class
