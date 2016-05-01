Attribute VB_Name = "Conexiones"

'Crea una string con la conexion para los reportes
Function Connect() As String

    Dim DbConnect$, DSN$, UID$, PWD$, DSQ$
    
    DbConnect = db.Connect
    DSN = getDsn(DbConnect)
    UID = getUid(DbConnect)
    PWD = getPwd(DbConnect)
    DSQ = getDatabase(DbConnect)
    
    'Connect = "DSN=" & DSN & ";" & "UID=" & UID & ";" & "DSN=" & DSN & ";" & "PWD=" & PWD & ";" & "DSQ=" & DSQ
    Connect = "DSN=" & DSN & ";" & "UID=" & UID & ";" & "PWD=" & PWD & ";" & "DSQ=" & DSQ
    
End Function
Function getPwd(Connect$) As String

    Dim i%, j%, aux$, largo%, Clave$
    
    Clave$ = "PWD="
    largo% = Len(Clave)
    aux = ""
    i = InStr(UCase(Connect), Clave) + largo
    
    If i <> 0 Then
        j = InStr(i, Connect, ";")
        If j <> 0 Then
            aux = Mid(Connect, i, j - i)
        Else
            aux = Mid(Connect, i, Len(Connect) - i + 1)
        End If
    End If
    
    getPwd = aux

End Function

Function getUid(Connect$) As String

    Dim i%, j%, aux$, largo%, Clave$
    
    Clave$ = "UID="
    largo% = Len(Clave)
    aux = ""
    i = InStr(UCase(Connect), Clave) + largo
    
    If i <> 0 Then
        j = InStr(i, Connect, ";")
        If j <> 0 Then
            aux = Mid(Connect, i, j - i)
        Else
            aux = Mid(Connect, i, Len(Connect) - i + 1)
        End If
    End If
    
    getUid = aux

End Function

Function getDatabase(Connect$) As String

    Dim i%, j%, aux$, largo%, Clave$
    
    Clave$ = "DATABASE="
    largo% = Len(Clave)
    aux = ""
    i = InStr(UCase(Connect), Clave) + largo
    
    If i <> largo Then
        j = InStr(i, Connect, ";")
        If j <> 0 Then
            aux = Mid(Connect, i, j - i)
        Else
            aux = Mid(Connect, i, Len(Connect) - i + 1)
        End If
    End If
    
    getDatabase = aux

End Function


Function getDsn(Connect$) As String

    Dim i%, j%, aux$, largo%, Clave$
    
    Clave$ = "DSN="
    largo% = Len(Clave)
    aux = ""
    i = InStr(UCase(Connect), Clave) + largo
    
    If i <> 0 Then
        j = InStr(i, Connect, ";")
        If j <> 0 Then
            aux = Mid(Connect, i, j - i)
        Else
            aux = Mid(Connect, i, Len(Connect) - i + 1)
        End If
    End If
    
    getDsn = aux
    
End Function

