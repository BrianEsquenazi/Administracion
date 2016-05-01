VERSION 5.00
Begin VB.Form PrgProce1 
   Caption         =   "Generacion de traspaso de datos"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
End
Attribute VB_Name = "PrgProce1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub proceso()

    On Error GoTo Error
    
    coderr = 0
    
    'borra los ensayos
    
    With rstWEnsayos
        .Index = "Codigo"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'borra los especificaciones
    
    With rstWEspecificaciones
        .Index = "Producto"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'borra las especificaciones
    
    With rstWEspecif
        .Index = "Producto"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'pruebas de laboratorio
    
    With rstWPrueba
        .Index = "Prueba"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'pruebas de laboratorio
    
    With rstWPrueter
        .Index = "Prueba"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'cotizaciones
    
    With rstWCotiza
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'hoja de produccion
    
    With rstWHoja
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'informe de produccion
    
    With rstWInforme
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'laudo de liberacion
    
    With rstWLaudo
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'Orden de compra
    
    With rstWOrden
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'mmovimientos varios de stock
    
    With rstWMovvar
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'materia prima
    
    With rstWArticulo
        .Index = "Codigo"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'Cliente
    
    With rstWClientes
        .Index = "Cliente"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'composicion de productos
    
    With rstWComposicion
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'cuenta corriente de clientes
    
    With rstWCtaCte
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'Leyendas
    
    With rstWDescComp
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'Precios
    
    With rstWPrecios
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'productos terminados
    
    With rstWTerminado
        .Index = "Codigo"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    'estadistica de ventas
    
    With rstWEstadistica
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
        
    'ensayos
        
    With rstEnsayos
            .Index = "Codigo"
            .MoveFirst
            Do
            
                If !Wdate = WFectraspaso Then
                
                    WEnsayo = !Codigo
                    WDescripcion = !Descripcion
                    
                    With rstWEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(WEnsayo)
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WEnsayo
                            !Descripcion = WDescri
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WEnsayo
                            !Descripcion = WDescri
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE MATERIAS PRIMAS
        
    With rstPrueba
            .Index = "Prueba"
            .MoveFirst
            Do
            
                If !Wdate = WFectraspaso Then
                
                    WPrueba = !Prueba
                    WProducto = !Producto
                    WFecha = !Fecha
                    Worden = !orden
                    WValor1 = !Valor1
                    WValor1 = !valor2
                    WValor1 = !Valor3
                    WValor1 = !valor4
                    WValor1 = !valor5
                    WValor1 = !valor6
                    WValor1 = !valor7
                    WValor1 = !valor8
                    WValor1 = !valor9
                    WValor1 = !valor10
                    WEnsayo = !Ensayo
                    WAspecto = !Aspecto
                    WObservaciones = !Observaciones
                    WConfecciono = !Confecciono
                    WLiberada = !Liberada
                    WDevuelta = !Devuelta
                    WLote = !Lote
                    WRechazo = !Rechazo
                    WNueva = !Nueva
                    
                    With rstWPrueba
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !orden = Worden
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote.Text
                            !Rechazo = WRechazo
                            !Nueva = WNueva
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !orden = Worden
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote.Text
                            !Rechazo = WRechazo
                            !Nueva = WNueva
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE productos terminados
        
    With rstPrueter
            .Index = "Prueba"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                    WPrueba = !Prueba
                    WProducto = !Producto
                    WFecha = !Fecha
                    WValor1 = !Valor1
                    WValor2 = !valor2
                    WValor3 = !Valor3
                    WValor4 = !valor4
                    WValor5 = !valor5
                    WValor6 = !valor6
                    WValor7 = !valor7
                    WValor8 = !valor8
                    WValor9 = !valor9
                    WValor10 = !valor10
                    WEnsayo = !Ensayo
                    WAspecto = !Aspecto
                    WObservaciones = !Observaciones
                    WConfecciono = !Confecciono
                    WLote = !Lote
                    WRechazo = !Rechazo
                    
                    With rstWPrueter
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Lote = WLote
                            !Rechazo = WRechazo
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Lote = WLote
                            !Rechazo = WRechazo
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    'especificaciones de p.t.
        
    With rstEspecif
            .Index = "Producto"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                    WProducto = !Producto
                    WEnsayo1 = !Ensayo1
                    WEnsayo2 = !Ensayo2
                    WEnsayo3 = !Ensayo3
                    WEnsayo4 = !Ensayo4
                    WEnsayo5 = !Ensayo5
                    WEnsayo6 = !Ensayo6
                    WEnsayo7 = !Ensayo7
                    WEnsayo8 = !Ensayo8
                    WEnsayo9 = !Ensayo9
                    WEnsayo10 = !Ensayo10
                    WValor1 = !Valor1
                    WValor2 = !valor2
                    WValor3 = !Valor3
                    WValor4 = !valor4
                    WValor5 = !valor5
                    WValor6 = !valor6
                    WValor7 = !valor7
                    WValor8 = !valor8
                    WValor9 = !valor9
                    WValor10 = !valor10
                    
                    With rstWEspecif
                        .Index = "Producto"
                        .Seek "=", WProducto
                        If .NoMatch Then
                            .AddNew
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                 End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With


    'especificaciones de m.p.
        
    With rstEspecificaciones
            .Index = "Producto"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                    WProducto = !Producto
                    WEnsayo1 = !Ensayo1
                    WEnsayo2 = !Ensayo2
                    WEnsayo3 = !Ensayo3
                    WEnsayo4 = !Ensayo4
                    WEnsayo5 = !Ensayo5
                    WEnsayo6 = !Ensayo6
                    WEnsayo7 = !Ensayo7
                    WEnsayo8 = !Ensayo8
                    WEnsayo9 = !Ensayo9
                    WEnsayo10 = !Ensayo10
                    WValor1 = !Valor1
                    WValor2 = !valor2
                    WValor3 = !Valor3
                    WValor4 = !valor4
                    WValor5 = !valor5
                    WValor6 = !valor6
                    WValor7 = !valor7
                    WValor8 = !valor8
                    WValor9 = !valor9
                    WValor10 = !valor10
                    
                    With rstWEspecificaciones
                        .Index = "Producto"
                        .Seek "=", WProducto
                        If .NoMatch Then
                            .AddNew
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With


    PrgProce1.Hide
    Unload Me
    traspa.SetFocus
    
Error:
     coderr = Err
     Resume Next
     
End Sub

Sub Form_Load()
    Call proceso
End Sub

