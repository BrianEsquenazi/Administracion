Public Class CustomConvert

    Public Shared Function toDoubleOrZero(ByVal value)
        Return toDoubleOr(value, 0)
    End Function

    Public Shared Function toDoubleOr(ByVal value, ByVal defaultValue)
        Try
            Return Convert.ToDouble(value)
        Catch
            Return defaultValue
        End Try
    End Function

    Public Shared Function toIntOrNull(ByVal value)
        Return toIntOr(value, Nothing)
    End Function

    Public Shared Function toIntOrZero(ByVal value)
        Return toIntOr(value, 0)
    End Function

    Public Shared Function toIntOr(ByVal value, ByVal defaultValue)
        Try
            Return Convert.ToInt32(value)
        Catch
            Return defaultValue
        End Try
    End Function

    Public Shared Function asTextDate(ByVal value)
        Try
            Dim myDate As Date = Convert.ToDateTime(value)
            Return myDate.ToShortDateString
        Catch ex As Exception
            Return "  /  /    "
        End Try
    End Function

End Class
