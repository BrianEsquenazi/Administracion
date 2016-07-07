Public Class QueryController
    Public text As String
    Public query As QueryFunction
    Public showMethod As ShowMethod

    Public Sub New(ByVal description As String, ByVal queryFunction As QueryFunction, ByVal showFunction As ShowMethod)
        text = description
        query = queryFunction
        showMethod = showFunction
    End Sub

    Public Overrides Function ToString() As String
        Return text
    End Function
End Class
