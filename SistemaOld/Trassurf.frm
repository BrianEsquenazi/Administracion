VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgtrassurf 
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Prgtrassurf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    WFectraspaso = Mid$(Fecha.Text, 4, 2) + "-" + Left$(Fecha.Text, 2) + "-" + Right$(Fecha.Text, 4)
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
    Prgtrassurf.Hide
    Unload Me
    Menu.SetFocus
End Sub


Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    Rem OPEN_FILE_Orden
    Rem OPEN_FILE_Cotiza
    Rem OPEN_FILE_Clientes
    
    OPEN_FILE_Articulo
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
    OPEN_FILE_Ctacte
    OPEN_FILE_DescComp
    OPEN_FILE_Precios
    OPEN_FILE_TERMINADO
    OPEN_FILE_Estadistica
    
    Rem OPEN_FILE_WOrden
    Rem OPEN_FILE_WCotiza
    Rem OPEN_FILE_WClientes
    
    OPEN_FILE_WArticulo
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
    OPEN_FILE_WCtacte
    OPEN_FILE_WDescComp
    OPEN_FILE_WPrecios
    OPEN_FILE_WTERMINADO
    OPEN_FILE_WEstadistica
    
    
    'borra los ensayos
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
    With rstWEspeci
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
    
    coderr = 0
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
    
    coderr = 0
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

    'movimientos varios de laboratorio
    
    coderr = 0
    With rstWMovlab
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
    
    'cotizaciones
    
    Rem coderr = 0
    Rem With rstWCotiza
    Rem     .Index = "Clave"
    Rem     .MoveFirst
    Rem     If coderr = 0 Then
    Rem         Do
    Rem             .Delete
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End If
    Rem End With
    
    'hoja de produccion
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
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
    
    Rem coderr = 0
    Rem With rstWOrden
    Rem     .Index = "Clave"
    Rem     .MoveFirst
    Rem     If coderr = 0 Then
    Rem         Do
    Rem             .Delete
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End If
    Rem End With
    
    'mmovimientos varios de stock
    
    coderr = 0
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
    
    coderr = 0
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
    
    Rem coderr = 0
    Rem With rstWClientes
    Rem     .Index = "Cliente"
    Rem     .MoveFirst
    Rem     If coderr = 0 Then
    Rem         Do
    Rem             .Delete
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End If
    Rem End With
    
    'composicion de productos
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
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
    
    coderr = 0
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
        
    coderr = 0
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE MATERIAS PRIMAS
        
    coderr = 0
    With rstPrueba
            .Index = "Prueba"
            .MoveFirst
            Do
            
                If !Wdate = WFectraspaso Then
                
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
                    
                    With rstWPrueba
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE productos terminados
        
    coderr = 0
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
        
    coderr = 0
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
        
    coderr = 0
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
    

    'moviienmtos varios de laboratiorio
        
    coderr = 0
    With rstMovlab
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                    
                    With rstWMovlab
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    

    Rem PRODUCTOS TERMINADOS
        
    coderr = 0
    With rstTerminado
            .Index = "Codigo"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                    
                    With rstWTerminado
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
   Rem clientes
        
    Rem coderr = 0
    Rem With rstClientes
    Rem        .Index = "Cliente"
    Rem         .MoveFirst
    Rem         Do
    Rem             If !Wdate = WFectraspaso Then
    Rem
    Rem                 WCliente = !Cliente
    Rem                 WRazon = !Razon
    Rem                 WDireccion = !Direccion
    Rem                 WLocalidad = !Localidad
    Rem                 WPostal = !Postal
    Rem                 WTelefono = !Telefono
    Rem                 WContacto = !Contacto
    Rem                 WObservaciones = !Observaciones
    Rem                 WCuit = !Cuit
    Rem                 WVendedor = !Vendedor
    Rem                 Wemail = !email
    Rem                 Wfax = !fax
    Rem                 WRubro = !Rubro
    Rem                 WHorario = !Horario
    Rem                 WPago1 = !Pago1
    Rem                 WPago2 = !pago2
    Rem                 WLimite = !Limite
    Rem                 WMinimo = !Minimo
    Rem                 WDirEntrega = !DirEntrega
    Rem                 WProvincia = !Provincia
    Rem                 WIva = !Iva
    Rem
    Rem                 With rstWClientes
    Rem                     .Index = "Cliente"
    Rem                     .Seek "=", WCliente
    Rem                     If .NoMatch Then
    Rem                         .AddNew
    Rem                         WCliente = !Cliente
    Rem                         WRazon = !Razon
    Rem                         WDireccion = !Direccion
    Rem                         WLocalidad = !Localidad
    Rem                         WPostal = !Postal
    Rem                         WTelefono = !Telefono
    Rem                         WContacto = !Contacto
    Rem                         WObservaciones = !Observaciones
    Rem                         WCuit = !Cuit
    Rem                         WVendedor = !Vendedor
    Rem                         Wemail = !email
    Rem                         Wfax = !fax
    Rem                         WRubro = !Rubro
    Rem                         WHorario = !Horario
    Rem                         WPago1 = !Pago1
    Rem                         WPago2 = !pago2
    Rem                         WLimite = !Limite
    Rem                         WMinimo = !Minimo
    Rem                         WDirEntrega = !DirEntrega
    Rem                         WProvincia = !Provincia
    Rem                         WIva = !Iva
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                             Else
    Rem                         .Edit
    Rem                         WCliente = !Cliente
    Rem                         WRazon = !Razon
    Rem                         WDireccion = !Direccion
    Rem                         WLocalidad = !Localidad
    Rem                         WPostal = !Postal
    Rem                         WTelefono = !Telefono
    Rem                         WContacto = !Contacto
    Rem                         WObservaciones = !Observaciones
    Rem                         WCuit = !Cuit
    Rem                         WVendedor = !Vendedor
    Rem                         Wemail = !email
    Rem                         Wfax = !fax
    Rem                         WRubro = !Rubro
    Rem                         WHorario = !Horario
    Rem                         WPago1 = !Pago1
    Rem                         WPago2 = !pago2
    Rem                         WLimite = !Limite
    Rem                         WMinimo = !Minimo
    Rem                         WDirEntrega = !DirEntrega
    Rem                         WProvincia = !Provincia
    Rem                         WIva = !Iva
    Rem                         .Update
    Rem                         .Bookmark = .LastModified
    Rem                     End If
    Rem                 End With
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    
   Rem precios por cliente
        
    coderr = 0
    With rstPrecios
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
                    WCliente = !Cliente
                    WTerminado = !Terminado
                    WPrecio = !Precio
                    WClave = !Clave
                    WDescripcion = !Descripcion
                    WFecha1 = !Fecha1
                    WFactura1 = !Factura1
                    WPrecio1 = !Precio1
                    WCantidad1 = !Cantidad1
                    WFecha2 = !Fecha2
                    WFactura2 = !Factura2
                    WPrecio2 = !Precio2
                    WCantidad2 = !Cantidad2
                    WFecha3 = !Fecha3
                    WFactura3 = !Factura3
                    WPrecio3 = !Precio3
                    WCantidad3 = !Cantidad3
                    WFecha4 = !Fecha4
                    WFactura4 = !Factura4
                    WPrecio4 = !Precio4
                    WCantidad4 = !Cantidad4
                    WFecha5 = !Fecha5
                    WFactura5 = !Factura5
                    WPrecio5 = !Precio5
                    WCantidad5 = !Cantidad5
                    
                    With rstWPrecios
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Cliente = WCliente
                            !Terminado = WTerminado
                            !Precio = WPrecio
                            !Clave = WClave
                            !Descripcion = WDescripcion
                            !Fecha1 = WFecha1
                            !Factura1 = WFactura1
                            !Precio1 = WPrecio1
                            !Cantidad1 = WCantidad1
                            !Fecha2 = WFecha2
                            !Factura2 = WFactura2
                            !Precio2 = WPrecio2
                            !Cantidad2 = WCantidad2
                            !Fecha3 = WFecha3
                            !Factura3 = WFactura3
                            !Precio3 = WPrecio3
                            !Cantidad3 = WCantidad3
                            !Fecha4 = WFecha4
                            !Factura4 = WFactura4
                            !Precio4 = WPrecio4
                            !Cantidad4 = WCantidad4
                            !Fecha5 = WFecha5
                            !Factura5 = WFactura5
                            !Precio5 = WPrecio5
                            !Cantidad5 = WCantidad5
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Cliente = WCliente
                            !Terminado = WTerminado
                            !Precio = WPrecio
                            !Clave = WClave
                            !Descripcion = WDescripcion
                            !Fecha1 = WFecha1
                            !Factura1 = WFactura1
                            !Precio1 = WPrecio1
                            !Cantidad1 = WCantidad1
                            !Fecha2 = WFecha2
                            !Factura2 = WFactura2
                            !Precio2 = WPrecio2
                            !Cantidad2 = WCantidad2
                            !Fecha3 = WFecha3
                            !Factura3 = WFactura3
                            !Precio3 = WPrecio3
                            !Cantidad3 = WCantidad3
                            !Fecha4 = WFecha4
                            !Factura4 = WFactura4
                            !Precio4 = WPrecio4
                            !Cantidad4 = WCantidad4
                            !Fecha5 = WFecha5
                            !Factura5 = WFactura5
                            !Precio5 = WPrecio5
                            !Cantidad5 = WCantidad5
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
    
   Rem Composicion
        
    coderr = 0
    With rstComposicion
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                                        
                    WTerminado = !Terminado
                    Wrenglon = !Renglon
                    WTipo = !Tipo
                    WArticulo1 = !Articulo1
                    WArticulo2 = !Articulo2
                    WCantidad = !Cantidad
                    WClave = !Clave
                    
                    With rstWComposicion
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem Articulo
        
    coderr = 0
    With rstArticulo
            .Index = "Codigo"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                    With rstWArticulo
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
                End If
    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem cuenta corriente
        
    coderr = 0
    With rstCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
                    WTipo = !Tipo
                    WNumero = !Numero
                    Wrenglon = !Renglon
                    WCliente = !Cliente
                    WFecha = !Fecha
                    WEstado = !Estado
                    Wvencimiento = !Vencimiento
                    WVencimiento1 = !Vencimiento1
                    WTotal = !Total
                    WTotalUs = !TotalUs
                    WSaldo = !Saldo
                    WSaldoUs = !SaldoUS
                    WOrdFecha = !OrdFecha
                    WOrdVencimiento = !OrdVencimiento
                    WOrdVencimiento1 = !OrdVencimiento1
                    WImpre = !Impre
                    WNeto = !Neto
                    WIva1 = !Iva1
                    WIva2 = !Iva2
                    WPedido = !Pedido
                    WRemito = !Remito
                    WOrden = !Orden
                    WParidad = !Paridad
                    WProvincia = !Provincia
                    WVendedor = !Vendedor
                    WRubro = !Rubro
                    WComprobante = !Comprobante
                    WAceptada = !Aceptada
                    WCosto = !Costo
                    WImporte1 = !Importe1
                    WImporte2 = !Importe2
                    WImporte3 = !Importe3
                    WImporte4 = !Importe4
                    WImporte5 = !Importe5
                    WImporte6 = !Importe6
                    WImporte7 = !Importe7
                    WClave = !Clave
                    
                    With rstWCtaCte
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
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
    
    
   Rem ESTADISTICA
        
    coderr = 0
    With rstEstadistica
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                    
                    With rstWEstadistica
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
   Rem Desfripcion de comprobantes
        
    coderr = 0
    With rstDescComp
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
                    WTipo = !Tipo
                    WNumero = !Numero
                    Wrenglon = !Renglon
                    WDescripcion = !Descripcion
                    WImporte = !Importe
                    WClave = !Clave
                    
                    With rstWDescComp
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Descripcion = WDescripcion
                            !Importe = WImporte
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Descripcion = WDescripcion
                            !Importe = WImporte
                            !Clave = WClave
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
    
    
    
   Rem cotizaciones
        
    Rem coderr = 0
    Rem With rstCotiza
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             If !Wdate = WFectraspaso Then
    Rem
    Rem                 WCotiza = !Cotiza
    Rem                 Wrenglon = !Renglon
    Rem                 WFecha = !Fecha
    Rem                 WProveedor = !Proveedor
    Rem                 WArticulo = !Articulo
    Rem                 WPrecio = !Precio
    Rem                 WFechaord = !FechaOrd
    Rem                 WCondicion = !Condicion
    Rem                 WObservaciones = !Observaciones
    Rem                 WClave = !Clave
    Rem
    Rem                 With rstWCotiza
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
    Rem                 End With
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
   Rem ORDENES DE COMPRA
        
    Rem coderr = 0
    Rem With rstOrden
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             If !Wdate = WFectraspaso Then
    Rem
    Rem                 WOrden = !Orden
    Rem                 Wrenglon = !Renglon
    Rem                 WFecha = !Fecha
    Rem                 WFechaord = !FechaOrd
    Rem                 WProveedor = !Proveedor
    Rem                 WArticulo = !Articulo
    Rem                 WCantidad = !Cantidad
    Rem                 WPrecio = !Precio
    Rem                 WFecha1 = !Fecha1
    Rem                 WFecha2 = !Fecha2
    Rem                 WCondicion = !Condicion
    Rem                 WRecibida = !Recibida
    Rem                 WClave = !Clave
    Rem
    Rem                 With rstWOrden
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
    Rem                 End With
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
   Rem INFORME DE RECEPCION
        
    coderr = 0
    With rstInforme
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                
                    With rstWInforme
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
    
    
   Rem LAUDO DE LIBERACION
        
    coderr = 0
    With rstLaudo
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                
                    With rstWLaudo
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
    
   Rem MOVIIMENTOS VARIOS
        
    coderr = 0
    With rstMovvar
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                
                    With rstWInforme
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    
   Rem HOJA DE PRODUCCION
        
    coderr = 0
    With rstHoja
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
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
                
                    With rstWHoja
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Rem With rstOrden
    Rem     .Close
    Rem End With
    Rem With rstCotiza
    Rem    .Close
    Rem End With
    Rem With rstClientes
    Rem    .Close
    Rem End With
    
    With rstArticulos
       .Close
    End With
    With rstEnsayos
        .Close
    End With
    With rstEspecificaciones
        .Close
    End With
    With rstEspecif
        .Close
    End With
    With rstPrueba
        .Close
    End With
    With rstPrueter
        .Close
    End With
    With rstMovlab
        .Close
    End With
    With rstHoja
        .Close
    End With
    With rstInforme
        .Close
    End With
    With rstLaudo
        .Close
    End With
    With rstMovvar
        .Close
    End With
    With rstComposicion
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    With rstDescomp
        .Close
    End With
    With rstPrecios
        .Close
    End With
    
    
    Rem With rstWCotiza
    Rem     .Close
    Rem End With
    Rem With rstWOrden
    Rem     .Close
    Rem End With
    Rem With rstWClientes
    Rem     .Close
    Rem End With
    
    With rstWArticulos
        .Close
    End With
    With rstWEnsayos
        .Close
    End With
    With rstWEspecificaciones
        .Close
    End With
    With rstWEspecif
        .Close
    End With
    With rstWPrueba
        .Close
    End With
    With rstWPrueter
        .Close
    End With
    With rstWMovlab
        .Close
    End With
    With rstWHoja
        .Close
    End With
    With rstWInforme
        .Close
    End With
    With rstWLaudo
        .Close
    End With
    With rstWMovvar
        .Close
    End With
    With rstWComposicion
        .Close
    End With
    With rstWCtaCte
        .Close
    End With
    With rstWDescomp
        .Close
    End With
    With rstWPrecios
        .Close
    End With
    
    DbsAdminis.Close
    DbsVentas.Close
    DbsCotiza.Close
    DbsLabora.Close
    DbsTraspa.Close
    
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
    
    'cotizaciones
    
    coderr = 0
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
    
    
    'Orden de compra
    
    coderr = 0
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
    
   Rem cotizaciones
        
    coderr = 0
    With rstCotiza
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
    
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
    
                    With rstWCotiza
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
                End If
    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem ORDENES DE COMPRA
        
    coderr = 0
    With rstOrden
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
    
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
    
                    With rstWOrden
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
                End If
        
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
    With rstOrden
        .Close
    End With
    With rstCotiza
       .Close
    End With
    
    With rstWCotiza
        .Close
    End With
    With rstWOrden
        .Close
    End With
    
    DbsCotiza.Close
    DbsOtro.Close

    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub




