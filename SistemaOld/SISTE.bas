Attribute VB_Name = "SISTEMA"
Global text As String


Sub NumbersOnly(T As Control, KeyAscii As Integer)
'This Sub allows only the digits 0 to 9, an initial minus sign and one period.
If KeyAscii < Asc(" ") Then     ' Is this Control char?
    Exit Sub                    ' Yes, let it pass
End If
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
     'don't discard it
ElseIf KeyAscii = Asc(".") Then 'if its a period
     If InStr(1, T, ".") Then 'if there is already a period
          KeyAscii = 0   'discard it
     End If
ElseIf KeyAscii = Asc("-") And T.SelStart = 0 Then
     'keep it, it's an initial minus sign
Else
    KeyAscii = 0  ' Discard all other characters
End If
'Now prevent any characters in front of a minus sign
If Mid$(T.text, T.SelStart + T.SelLength + 1, 1) = "-" Then
    KeyAscii = 0   ' Discard characters before -
End If
End Sub

Sub Errores(coderr As Integer, Archivo As String, Mensaje As String)

    e = coderr
    Select Case e
        Case 3021
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de " + Archivo$)
        Case Else
            M$ = Mensaje$
            A% = MsgBox(M$, 0, "Archivo de Ensayos")
    End Select
    
End Sub

Sub Ceros(Campo As String, Largo As Integer)

    L% = 1
    cadena$ = ""
    While L% <= Len(Campo) And L% > 0
        If Mid$(Campo, L%, 1) <> Chr$(32) Then cadena$ = cadena$ + Mid$(Campo, L%, 1)
        L% = L% + 1
    Wend
    Campo = Right$(String$(40, "0") + cadena$, Largo)
    
End Sub

