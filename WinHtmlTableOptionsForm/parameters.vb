Public Class parameters
    Private _FileName As String
    Private _TableCaption As String

    Public Property FileName() As String
        Get
            Return Me._FileName
        End Get
        Set(ByVal value As String)
            Me._FileName = value
        End Set
    End Property

    Public Property TableCaption() As String
        Get
            Return Me._TableCaption
        End Get
        Set(ByVal value As String)
            Me._TableCaption = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal FileName As String, ByVal TableCaption As String)
        Me._FileName = FileName
        Me._TableCaption = TableCaption
    End Sub
End Class
