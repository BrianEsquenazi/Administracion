﻿Public Class ProcesoReteIbRecibos

    Public fechaord As String
    Public retotra As Double
    Public renglon As String
    Public cliente As String
    Public fecha As String
    Public comproib As String
    Public recibo As String
    Public cuit As String

    Public Sub New(ByVal fechaord2 As String, ByVal retotra2 As Double, ByVal renglon2 As String, ByVal cliente2 As String,
                   ByVal fecha2 As String, ByVal comproib2 As String, ByVal cuit2 As String)

        fechaord = fechaord2
        retotra = retotra2
        renglon = renglon2
        cliente = cliente2
        fecha = fecha2
        comproib = comproib2
        REM recibo = recibo2
        cuit = cuit2

    End Sub

End Class
