Imports RestSharp

Friend Class RestClient
    Private v As String

    Public Sub New(v As String)
        Me.v = v
    End Sub

    Friend Sub Execute(request As RestRequest)
        Throw New NotImplementedException()
    End Sub
End Class
