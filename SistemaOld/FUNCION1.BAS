Attribute VB_Name = "Funciones"
Function PUsing(Mask As String, text As String)

         Dim T As Single
         Dim T1 As Single
         Dim Tx As String
         Dim T1x As String
         Dim Places As Single
         Dim Auxi As String
         
         For T = 1 To Len(text)
            If Mid$(text, T, 1) <> "," Then
                Auxi = Auxi + Mid$(text, T, 1)
            End If
         Next T
         
         text = Auxi
         
         x# = Val(text)
         PUsing = text

         T = Len(Mask)
         If T > 24 Then Error 5
         Tx = Space$(T)

         Places = 0
         T = InStr(Mask, ".")
        If T Then Places = Len(Mask) - T

         T1x = Mid$(Str$(Int(Abs(x#) + _
          (0.5 / 10# ^ Places))), 2 - _
          (Abs(x#) < 1#)) + _
          Mid$(".", 2 + (Places > 0)) + _
          Right$(Str$(Int((Abs(x#) + _
          10# ^ Places) * 10# ^ Places + _
          0.5)), Places)

         If Left$(T1x, 1) = "," Then T1x = "0" + T1x
         If InStr(Mask, "$") Then T1x = "$" + T1x
         If Sgn(x#) = -1 Then T1x = "-" + T1x

         If InStr(Mask, "+") And _
                 Left$(T1x, 1) <> "-" Then _
                 T1x = "+" + T1x

         If InStr(Mask, ",") Then
                 T = InStr(T1x, ".")
                 If T = 0 Then T = Len(T1x) + 1

                 For T = T - 4 To 1 Step -3
                T1 = Asc(Mid$(T1x, T))
                If T1 > 47 And T1 < 58 Then _
                        T1x = Left$(T1x, T) + "," + _
                         Mid$(T1x, T + 1)
       Next
    End If

         If Len(T1x) > Len(Tx) Then
                 Tx = "%" + T1x
         Else
                 RSet Tx = T1x
    End If

         If InStr(Mask, "*") Then
       For T = 1 To Len(Tx)
           If Mid$(Tx, T, 1) = " " Then Mid$(Tx, T, 1) = "*"
       Next
    End If
    
    PUsing = Trim(Tx)

End Function

Sub Valida_fecha(fecha As String, Validate As String)

    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    Dia = Left$(fecha, 2)
    Mes = Mid$(fecha, 4, 2)
    Ano = Right$(fecha, 4)
    
    Validate = "S"
    
    If Val(Dia) <= 0 Or Val(Dia) > 31 Then
        Validate = "N"
    End If
    
    If Val(Mes) <= 0 Or Val(Mes) > 12 Then
        Validate = "N"
    End If
    
    If Val(Ano) <= 1900 Then
        Validate = "N"
    End If

End Sub


Sub Verifica_datos(Dato As String, Valores As String, Valida As String)

    Dim Largo1 As Single
    Dim Largo2 As Single
    Dim Cicla1 As Single
    Dim Cicla2 As Single
    
    Largo1 = Len(Dato)
    Largo2 = Len(Valores)
    Valida = "N"
    
    For Cicla2 = 1 To Largo2
        If Dato = Mid$(Valores, Cicla2, 1) Then
            Valida = "S"
            Exit For
        End If
    Next Cicla2

End Sub





