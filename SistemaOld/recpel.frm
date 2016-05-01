VERSION 5.00
Begin VB.Form PrgRecpel 
   Caption         =   "Generacion de traspaso de datos"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "PrgRecpel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    WEmpresa = "0001"
    Call Proceso
    WEmpresa = "0002"
    Call Proceso
    WEmpresa = "0003"
    Call Proceso1
    WEmpresa = "0004"
    Call Proceso1
    WEmpresa = "0005"
    Call Proceso1
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgRecsurf.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    OPEN_FILE_ENSAYOS
    OPEN_FILE_ESPECIFICACIONES
    OPEN_FILE_ESPECIF
    OPEN_FILE_PRUEBA
    OPEN_FILE_PrueTer
    OPEN_FILE_Movlab
    OPEN_FILE_Hoja
    OPEN_FILE_Informe
    OPEN_FILE_LAUDO
    OPEN_FILE_Movvar
    OPEN_FILE_Composicion
    OPEN_FILE_TERMINADO
    OPEN_FILE_Estadistica
    OPEN_FILE_Articulo
    
    OPEN_FILE_WENSAYOS
    OPEN_FILE_WESPECIFICACIONES
    OPEN_FILE_WESPECIF
    OPEN_FILE_WPRUEBA
    OPEN_FILE_WPrueTer
    OPEN_FILE_WMovlab
    OPEN_FILE_WHoja
    OPEN_FILE_WInforme
    OPEN_FILE_WLAUDO
    OPEN_FILE_WMovvar
    OPEN_FILE_WComposicion
    OPEN_FILE_WTERMINADO
    OPEN_FILE_WEstadistica
    OPEN_FILE_WArticulo
    
    'ensayos
        
    coderr = 0
    With rstWEnsayos
            .Index = "Codigo"
            .MoveFirst
            Do
            
                WEnsayo = !Codigo
                WDescripcion = !Descripcion
                    
                With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(WEnsayo)
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE MATERIAS PRIMAS
        
    coderr = 0
    With rstWPrueba
            .Index = "Prueba"
            .MoveFirst
            Do
            
                WPrueba = !Prueba
                WProducto = !Producto
                WFecha = !Fecha
                WOrden = !Orden
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
                WLiberada = !Liberada
                WDevuelta = !Devuelta
                WLote = !Lote
                WRechazo = !Rechazo
                WNueva = !Nueva
                
                With rstPrueba
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Orden = WOrden
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
                            !Orden = WOrden
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
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE productos terminados
        
    coderr = 0
    With rstWPrueter
            .Index = "Prueba"
            .MoveFirst
            Do
            
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
                
                With rstPrueter
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
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    'especificaciones de p.t.
        
    coderr = 0
    With rstWEspecif
            .Index = "Producto"
            .MoveFirst
            Do
            
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
                
                With rstEspecif
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
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With


    'especificaciones de m.p.
        
    coderr = 0
    With rstWEspecificaciones
            .Index = "Producto"
            .MoveFirst
            Do
            
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
                    
                With rstEspecificaciones
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
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    

    'moviienmtos varios de laboratiorio
        
    coderr = 0
    With rstWMovlab
            .Index = "Clave"
            .MoveFirst
            Do
                
                WCodigo = !Codigo
                Wrenglon = !Renglon
                WFecha = !Fecha
                WFechaord = !FechaOrd
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = !Cantidad
                WMovi = !Movi
                WTipomov = !Tipomov
                WObservaciones = !Observaciones
                WClave = !Clave
                
                With rstMovlab
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    

    Rem PRODUCTOS TERMINADOS
        
    coderr = 0
    With rstWTerminado
            .Index = "Codigo"
            .MoveFirst
            Do
                
                WCodigo = !Codigo
                WDescripcion = !Descripcion
                WLinea = !Linea
                WUnidad = !Unidad
                WInicial = !Inicial
                WEntradas = !Entradas
                WSalidas = !Salidas
                WMinimo = !Minimo
                WDeposito = !Deposito
                WPedido = !Pedido
                WEnvase1 = !Envase1
                WEnvase2 = !Envase2
                WEnvase3 = !Envase3
                WEnvase4 = !Envase4
                WEnvase5 = !Envase5
                WEnvase6 = !Envase6
                WProceso = !Proceso
                WCosto = !Costo
                WFactor = !Factor
                
                With rstTerminado
                        .Index = "Codigo"
                        .Seek "=", WCodigo
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Linea = WLinea
                            !Unidad = WUnidad
                            !Inicial = WInicial
                            !Entradas = WEntradas
                            !Salidas = WSalidas
                            !Minimo = WMinimo
                            !Deposito = WDeposito
                            !Pedido = WPedido
                            !Envase1 = WEnvase1
                            !Envase2 = WEnvase2
                            !Envase3 = WEnvase3
                            !Envase4 = WEnvase4
                            !Envase5 = WEnvase5
                            !Envase6 = WEnvase6
                            !Proceso = WProceso
                            !Costo = WCosto
                            !Factor = WFactor
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Linea = WLinea
                            !Unidad = WUnidad
                            !Inicial = WInicial
                            !Entradas = WEntradas
                            !Salidas = WSalidas
                            !Minimo = WMinimo
                            !Deposito = WDeposito
                            !Pedido = WPedido
                            !Envase1 = WEnvase1
                            !Envase2 = WEnvase2
                            !Envase3 = WEnvase3
                            !Envase4 = WEnvase4
                            !Envase5 = WEnvase5
                            !Envase6 = WEnvase6
                            !Proceso = WProceso
                            !Costo = WCosto
                            !Factor = WFactor
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
   Rem Composicion
        
    coderr = 0
    With rstWComposicion
            .Index = "Clave"
            .MoveFirst
            Do
                                        
                WTerminado = !Terminado
                Wrenglon = !Renglon
                WTipo = !Tipo
                WArticulo1 = !Articulo1
                WArticulo2 = !Articulo2
                WCantidad = !Cantidad
                WClave = !Clave
                    
                With rstComposicion
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Terminado = WTerminado
                            !Renglon = Wrenglon
                            !Tipo = WTipo
                            !Articulo1 = WArticulo1
                            !Articulo2 = WArticulo2
                            !Cantidad = WCantidad
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Terminado = WTerminado
                            !Renglon = Wrenglon
                            !Tipo = WTipo
                            !Articulo1 = WArticulo1
                            !Articulo2 = WArticulo2
                            !Cantidad = WCantidad
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem Articulo
        
    coderr = 0
    With rstWArticulo
            .Index = "Codigo"
            .MoveFirst
            Do
    
                WCodigo = !Codigo
                WDescripcion = !Descripcion
                WCosto1 = !Costo1
                WCosto2 = !Costo2
                WInicial = !Inicial
                WEntradas = !Entradas
                WSalidas = !Salidas
                WMinimo = !Minimo
                WLaboratorio = !Laboratorio
                WUnidad = !Unidad
                WPedido = !Pedido
                WDeposito = !Deposito
                WEnvase = !Envase
                WRs = !Rs
                WProveedor = !Proveedor
                WFecha = !Fecha
    
                With rsWArticulo
                        .Index = "Codigo"
                        .Seek "=", WCodigo
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Inicial = WInicial
                            !Entradas = WEntradas
                            !Salidas = WSalidas
                            !Minimo = WMinimo
                            !Laboratorio = WLaboratorio
                            !Unidad = WUnidad
                            !Pedido = WPedido
                            !Deposito = WDeposito
                            !Envase = WEnvase
                            !Rs = WRs
                            !Proveedor = WProveedor
                            !Fecha = WFecha
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Inicial = WInicial
                            !Entradas = WEntradas
                            !Salidas = WSalidas
                            !Minimo = WMinimo
                            !Laboratorio = WLaboratorio
                            !Unidad = WUnidad
                            !Pedido = WPedido
                            !Deposito = WDeposito
                            !Envase = WEnvase
                            !Rs = WRs
                            !Proveedor = WProveedor
                            !Fecha = WFecha
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem ESTADISTICA
        
    coderr = 0
    With rstWEstadistica
            .Index = "Clave"
            .MoveFirst
            Do
                
                WTipo = !Tipo
                WNumero = !Numero
                Wrenglon = !Renglon
                WArticulo = !Articulo
                WCantidad = !Cantidad
                WPrecio = !Precio
                WPrecioUs = !PrecioUs
                WImporte = !Importe
                WImporteUs = !ImporteUs
                WCliente = !Cliente
                WParidad = !Paridad
                WVendedor = !Vendedor
                WRubro = !Rubro
                WLinea = !Linea
                WCosto1 = !Costo1
                WCosto2 = !Costo2
                WCoeficiente = !Coeficiente
                WPedido = !Pedido
                WFecha = !Fecha
                WImporte1 = !Importe1
                WImporte2 = !Importe2
                WImporte3 = !Importe3
                WImporte4 = !Importe4
                WOrdFecha = !OrdFecha
                WWArticulo = !WArticulo
                WRemito = !Remito
                WClave = !Clave
                
                With rstEstadistica
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WImporteUs
                            !Cliente = WCliente
                            !Paridad = WParidad
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Linea = WLinea
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Coeficiente = WCoeficiente
                            !Pedido = WPedido
                            !Fecha = WFecha
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !OrdFecha = WOrdFecha
                            !WArticulo = WWArticulo
                            !Remito = WRemito
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WImporteUs
                            !Cliente = WCliente
                            !Paridad = WParidad
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Linea = WLinea
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Coeficiente = WCoeficiente
                            !Pedido = WPedido
                            !Fecha = WFecha
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !OrdFecha = WOrdFecha
                            !WArticulo = WWArticulo
                            !Remito = WRemito
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
   Rem cotizaciones
        
    Rem coderr = 0
    Rem With rstWCotiza
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem
    Rem             WCotiza = !Cotiza
    Rem             Wrenglon = !Renglon
    Rem             WFecha = !Fecha
    Rem             WProveedor = !Proveedor
    Rem             WArticulo = !Articulo
    Rem             WPrecio = !Precio
    Rem             WFechaord = !FechaOrd
    Rem             WCondicion = !Condicion
    Rem             WObservaciones = !Observaciones
    Rem             WClave = !Clave
    Rem
    Rem             With rstCotiza
    Rem                     .Index = "Clave"
    Rem                     .Seek "=", WClave
    Rem                     If .NoMatch Then
    Rem                         .AddNew
    Rem                         !Cotiza = WCotiza
    Rem                         !Renglon = Wrenglon
    Rem                         !Fecha = WFecha
    Rem                         !Proveedor = WProveedor
    Rem                         !Articulo = WArticulo
    Rem                         !Precio = WPrecio
    Rem                         !FechaOrd = WFechaord
    Rem                         !Condicion = WCondicion
    Rem                         !Observaciones = WObservaciones
    Rem                         !Clave = WClave
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                             Else
    Rem                         .Edit
    Rem                         !Cotiza = WCotiza
    Rem                         !Renglon = Wrenglon
    Rem                         !Fecha = WFecha
    Rem                         !Proveedor = WProveedor
    Rem                         !Articulo = WArticulo
    Rem                         !Precio = WPrecio
    Rem                         !FechaOrd = WFechaord
    Rem                         !Condicion = WCondicion
    Rem                         !Observaciones = WObservaciones
    Rem                         !Clave = WClave
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                     End If
    Rem             End With
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    
    
   Rem ORDENES DE COMPRA
        
    Rem coderr = 0
    Rem With rstWOrden
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem
    Rem             WOrden = !Orden
    Rem             Wrenglon = !Renglon
    Rem             WFecha = !Fecha
    Rem             WFechaord = !FechaOrd
    Rem             WProveedor = !Proveedor
    Rem             WArticulo = !Articulo
    Rem             WCantidad = !Cantidad
    Rem             WPrecio = !Precio
    Rem             WFecha1 = !Fecha1
    Rem             WFecha2 = !Fecha2
    Rem             WCondicion = !Condicion
    Rem             WRecibida = !Recibida
    Rem             WClave = !Clave
    Rem
    Rem             With rstOrden
    Rem                     .Index = "Clave"
    Rem                     .Seek "=", WClave
    Rem                     If .NoMatch Then
    Rem                         .AddNew
    Rem                         !Orden = WOrden
    Rem                         !Renglon = Wrenglon
    Rem                         !Fecha = WFecha
    Rem                         !FechaOrd = WFechaord
    Rem                         !Proveedor = WProveedor
    Rem                         !Articulo = WArticulo
    Rem                         !Cantidad = WCantidad
    Rem                         !Precio = WPrecio
    Rem                         !Fecha1 = WFecha1
    Rem                         !Fecha2 = WFecha2
    Rem                         !Condicion = WCondicion
    Rem                         !Recibida = WRecibida
    Rem                         !Clave = WClave
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                             Else
    Rem                         .Edit
    Rem                         !Orden = WOrden
    Rem                         !Renglon = Wrenglon
    Rem                         !Fecha = WFecha
    Rem                         !FechaOrd = WFechaord
    Rem                         !Proveedor = WProveedor
    Rem                         !Articulo = WArticulo
    Rem                         !Cantidad = WCantidad
    Rem                         !Precio = WPrecio
    Rem                         !Fecha1 = WFecha1
    Rem                         !Fecha2 = WFecha2
    Rem                         !Condicion = WCondicion
    Rem                         !Recibida = WRecibida
    Rem                         !Clave = WClave
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                     End If
    Rem             End With
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    
    
    
   Rem INFORME DE RECEPCION
        
    coderr = 0
    With rstWInforme
            .Index = "Clave"
            .MoveFirst
            Do
                
                WInforme = !Informe
                Wrenglon = !Renglon
                WFecha = !Fecha
                WProveedor = !Proveedor
                WRemito = !Remito
                WOrden = !Orden
                WArticulo = !Articulo
                WCantidad = !Cantidad
                WResta = !Resta
                WClave = !Clave
                WFechaord = !FechaOrd
            
                With rstInforme
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Informe = WInforme
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !FechaOrd = WFechaord
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Informe = WInforme
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !FechaOrd = WFechaord
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
    
    
   Rem LAUDO DE LIBERACION
        
    coderr = 0
    With rstWLaudo
            .Index = "Clave"
            .MoveFirst
            Do
                                    
                WLaudo = !Laudo
                Wrenglon = !Renglon
                WFecha = !Fecha
                WOrden = !Orden
                WArticulo = !Articulo
                WLiberada = !Liberada
                WDevuelta = !Devuelta
                WLiberada = !Liberada
                WLote = !Lote
                WRechazo = !Rechazo
                WActualiza = !Actualiza
                WMarca = !Marca
                WInforme = !Informe
                WClave = !Clave
                
                With rstLaudo
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Laudo = WLaudo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Liberada = WLiberada
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !Actualiza = WNuevo
                            !Marca = WMarca
                            !Informe = WInforme
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Laudo = WLaudo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Liberada = WLiberada
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !Actualiza = WNuevo
                            !Marca = WMarca
                            !Informe = WInforme
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
    
   Rem MOVIIMENTOS VARIOS
        
    coderr = 0
    With rstWMovvar
            .Index = "Clave"
            .MoveFirst
            Do
                
                WCodigo = !Codigo
                Wrenglon = !Renglon
                WFecha = !Fecha
                WFechaord = !FechaOrd
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = !Cantidad
                WMovi = !Movi
                WTipomov = !Tipomov
                WObservaciones = !Observaciones
                WClave = !Clave
                
                With rstInforme
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
   Rem HOJA DE PRODUCCION
        
    coderr = 0
    With rstWHoja
            .Index = "Clave"
            .MoveFirst
            Do
                
                WHoja = !Hoja
                Wrenglon = !Renglon
                WFecha = !Fecha
                WProducto = !Producto
                WTeorico = !Teorico
                WReal = !Real
                WfechaIng = !fechaIng
                WfechaIngord = !WfechaIngord
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = !Cantidad
                WLote = !Lote
                WClave = !Clave
                
                With rstHoja
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Hoja = WHoja
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WfechaIng
                            !fechaIngord = WWfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Hoja = WHoja
                            !Renglon = Wrenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WfechaIng
                            !fechaIngord = WWfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub


Private Sub Proceso1()

    On Error GoTo Error
    
    OPEN_FILE_Orden
    OPEN_FILE_Cotiza
    
    OPEN_FILE_WOrden
    OPEN_FILE_WCotiza
    
    
   Rem cotizaciones
        
    coderr = 0
    With rstWCotiza
            .Index = "Clave"
             .MoveFirst
             Do
    
                 WCotiza = !Cotiza
                 Wrenglon = !Renglon
                 WFecha = !Fecha
                 WProveedor = !Proveedor
                 WArticulo = !Articulo
                 WPrecio = !Precio
                 WFechaord = !FechaOrd
                 WCondicion = !Condicion
                 WObservaciones = !Observaciones
                 WClave = !Clave
    
                 With rstCotiza
                         .Index = "Clave"
                         .Seek "=", WClave
                         If .NoMatch Then
                             .AddNew
                             !Cotiza = WCotiza
                             !Renglon = Wrenglon
                             !Fecha = WFecha
                             !Proveedor = WProveedor
                             !Articulo = WArticulo
                             !Precio = WPrecio
                             !FechaOrd = WFechaord
                             !Condicion = WCondicion
                             !Observaciones = WObservaciones
                             !Clave = WClave
                             .Update
                             .Bookmark = .LastModified
                                 Else
                             .Edit
                             !Cotiza = WCotiza
                             !Renglon = Wrenglon
                             !Fecha = WFecha
                             !Proveedor = WProveedor
                             !Articulo = WArticulo
                             !Precio = WPrecio
                             !FechaOrd = WFechaord
                             !Condicion = WCondicion
                             !Observaciones = WObservaciones
                             !Clave = WClave
                             .Update
                             .Bookmark = .LastModified
                         End If
                 End With
    
                 .MoveNext
                 If .EOF = True Then
                     Exit Do
                 End If
             Loop
    End With
    
    
    
   Rem ORDENES DE COMPRA
        
    coderr = 0
    With rstWOrden
             .Index = "Clave"
             .MoveFirst
             Do
    
                 WOrden = !Orden
                 Wrenglon = !Renglon
                 WFecha = !Fecha
                 WFechaord = !FechaOrd
                 WProveedor = !Proveedor
                 WArticulo = !Articulo
                 WCantidad = !Cantidad
                 WPrecio = !Precio
                 WFecha1 = !Fecha1
                 WFecha2 = !Fecha2
                 WCondicion = !Condicion
                 WRecibida = !Recibida
                 WClave = !Clave
    
                 With rstOrden
                         .Index = "Clave"
                         .Seek "=", WClave
                         If .NoMatch Then
                             .AddNew
                             !Orden = WOrden
                             !Renglon = Wrenglon
                             !Fecha = WFecha
                             !FechaOrd = WFechaord
                             !Proveedor = WProveedor
                             !Articulo = WArticulo
                             !Cantidad = WCantidad
                             !Precio = WPrecio
                             !Fecha1 = WFecha1
                             !Fecha2 = WFecha2
                             !Condicion = WCondicion
                             !Recibida = WRecibida
                             !Clave = WClave
                             .Update
                             .Bookmark = .LastModified
                                 Else
                             .Edit
                             !Orden = WOrden
                             !Renglon = Wrenglon
                             !Fecha = WFecha
                             !FechaOrd = WFechaord
                             !Proveedor = WProveedor
                             !Articulo = WArticulo
                             !Cantidad = WCantidad
                             !Precio = WPrecio
                             !Fecha1 = WFecha1
                             !Fecha2 = WFecha2
                             !Condicion = WCondicion
                             !Recibida = WRecibida
                             !Clave = WClave
                             .Update
                             .Bookmark = .LastModified
                         End If
                 End With
    
                 .MoveNext
                 If .EOF = True Then
                     Exit Do
                 End If
             Loop
    End With
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub





