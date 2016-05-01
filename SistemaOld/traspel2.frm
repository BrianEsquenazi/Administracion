VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgtraspel2 
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
Attribute VB_Name = "Prgtraspel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstPrueart As Recordset
Dim spPrueart As String
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEspecif As Recordset
Dim spEspecif As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim rstMovlab As Recordset
Dim spmovlab As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstDesccomp As Recordset
Dim spDesccomp As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim XParam As String
Dim A1 As String
Dim A2 As String

Private Sub Acepta_Click()

    WFectraspaso = Mid$(Fecha.Text, 4, 2) + "-" + Left$(Fecha.Text, 2) + "-" + Right$(Fecha.Text, 4)
    
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    WEmpresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    WEmpresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    WEmpresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prgtraspel2.Hide
    Unload Me
    End
End Sub


Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Call Acepta_Click
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
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
    OPEN_FILE_WMovguia
    OPEN_FILE_WClientes
    OPEN_FILE_WComposicion
    OPEN_FILE_WCtacte
    OPEN_FILE_WDescComp
    OPEN_FILE_WPrecios
    OPEN_FILE_WTERMINADO
    OPEN_FILE_WEstadistica
    OPEN_FILE_WSolic
    
    
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
    
    
    'mmovimientos varios de stock
    
    coderr = 0
    With rstWMovguia
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
    
    coderr = 0
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
    
    'solicitud de ordenes de compra
    
    coderr = 0
    With rstWSolic
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
        
    spEnsayo = "ListaEnsayos"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        
    coderr = 0
    With rstEnsayo
    
            .MoveFirst
            Do
            
                If rstEnsayo!WDate = WFectraspaso Then
                
                    WEnsayo = rstEnsayo!Codigo
                    WDescripcion = rstEnsayo!Descripcion
                    
                    With rstWEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(WEnsayo)
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !WDate = Date$
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !WDate = Date$
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
    rstEnsayo.Close
    
    End If

    
    'PRUEBAS DE MATERIAS PRIMAS
        
    spPrueart = "ListaPrueba"
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        
    coderr = 0
    With rstPrueart
            .MoveFirst
            Do
            
                If rstPrueart!WDate = WFectraspaso Then
                
                    WPrueba = rstPrueart!Prueba
                    WProducto = rstPrueart!Producto
                    WFecha = rstPrueart!Fecha
                    WOrden = rstPrueart!Orden
                    WValor1 = rstPrueart!Valor1
                    Wvalor2 = rstPrueart!valor2
                    WValor3 = rstPrueart!Valor3
                    Wvalor4 = rstPrueart!valor4
                    Wvalor5 = rstPrueart!valor5
                    Wvalor6 = rstPrueart!valor6
                    Wvalor7 = rstPrueart!valor7
                    Wvalor8 = rstPrueart!valor8
                    Wvalor9 = rstPrueart!valor9
                    Wvalor10 = rstPrueart!valor10
                    WEnsayo = rstPrueart!Ensayo
                    WAspecto = rstPrueart!Aspecto
                    WObservaciones = rstPrueart!Observaciones
                    WConfecciono = rstPrueart!Confecciono
                    WLiberada = rstPrueart!Liberada
                    WDevuelta = rstPrueart!Devuelta
                    WLote = rstPrueart!Lote
                    WRechazo = rstPrueart!Rechazo
                    WNueva = rstPrueart!Nueva
                    
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote
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
    
    rstPrueart.Close
    
    End If
    
    'PRUEBAS DE productos terminados
        
    spPrueter = "ListaPrueter"
    Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueter.RecordCount > 0 Then
        
    coderr = 0
    With rstPrueter
            .MoveFirst
            Do
                If rstPrueter!WDate = WFectraspaso Then
                    WPrueba = rstPrueter!Prueba
                    WProducto = rstPrueter!Producto
                    WFecha = rstPrueter!Fecha
                    WValor1 = rstPrueter!Valor1
                    Wvalor2 = rstPrueter!valor2
                    WValor3 = rstPrueter!Valor3
                    Wvalor4 = rstPrueter!valor4
                    Wvalor5 = rstPrueter!valor5
                    Wvalor6 = rstPrueter!valor6
                    Wvalor7 = rstPrueter!valor7
                    Wvalor8 = rstPrueter!valor8
                    Wvalor9 = rstPrueter!valor9
                    Wvalor10 = rstPrueter!valor10
                    WEnsayo = rstPrueter!Ensayo
                    WAspecto = rstPrueter!Aspecto
                    WObservaciones = rstPrueter!Observaciones
                    WConfecciono = rstPrueter!Confecciono
                    WLote = rstPrueter!Lote
                    WRechazo = rstPrueter!Rechazo
                    
                    With rstWPrueter
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Valor1 = WValor1
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
    
    rstPrueter.Close
    
    End If
    
    'especificaciones de p.t.
        
    spEspecif = "ListaEspecif"
    Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecif.RecordCount > 0 Then
        
    coderr = 0
    With rstEspecif
            .MoveFirst
            Do
                If rstEspecif!WDate = WFectraspaso Then
                    WProducto = rstEspecif!Producto
                    WEnsayo1 = rstEspecif!Ensayo1
                    WEnsayo2 = rstEspecif!Ensayo2
                    WEnsayo3 = rstEspecif!Ensayo3
                    WEnsayo4 = rstEspecif!Ensayo4
                    WEnsayo5 = rstEspecif!Ensayo5
                    WEnsayo6 = rstEspecif!Ensayo6
                    WEnsayo7 = rstEspecif!Ensayo7
                    WEnsayo8 = rstEspecif!Ensayo8
                    WEnsayo9 = rstEspecif!Ensayo9
                    WEnsayo10 = rstEspecif!Ensayo10
                    WValor1 = rstEspecif!Valor1
                    Wvalor2 = rstEspecif!valor2
                    WValor3 = rstEspecif!Valor3
                    Wvalor4 = rstEspecif!valor4
                    Wvalor5 = rstEspecif!valor5
                    Wvalor6 = rstEspecif!valor6
                    Wvalor7 = rstEspecif!valor7
                    Wvalor8 = rstEspecif!valor8
                    Wvalor9 = rstEspecif!valor9
                    Wvalor10 = rstEspecif!valor10
                    
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
    
    rstEspecif.Close
    
    End If

    'especificaciones de m.p.
        
    spEspecificaciones = "ListaEspecificaciones"
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificaciones.RecordCount > 0 Then
        
    coderr = 0
    With rstEspecificaciones
            .MoveFirst
            Do
                If rstEspecificaciones!WDate = WFectraspaso Then
                    WProducto = rstEspecificaciones!Producto
                    WEnsayo1 = rstEspecificaciones!Ensayo1
                    WEnsayo2 = rstEspecificaciones!Ensayo2
                    WEnsayo3 = rstEspecificaciones!Ensayo3
                    WEnsayo4 = rstEspecificaciones!Ensayo4
                    WEnsayo5 = rstEspecificaciones!Ensayo5
                    WEnsayo6 = rstEspecificaciones!Ensayo6
                    WEnsayo7 = rstEspecificaciones!Ensayo7
                    WEnsayo8 = rstEspecificaciones!Ensayo8
                    WEnsayo9 = rstEspecificaciones!Ensayo9
                    WEnsayo10 = rstEspecificaciones!Ensayo10
                    WValor1 = rstEspecificaciones!Valor1
                    Wvalor2 = rstEspecificaciones!valor2
                    WValor3 = rstEspecificaciones!Valor3
                    Wvalor4 = rstEspecificaciones!valor4
                    Wvalor5 = rstEspecificaciones!valor5
                    Wvalor6 = rstEspecificaciones!valor6
                    Wvalor7 = rstEspecificaciones!valor7
                    Wvalor8 = rstEspecificaciones!valor8
                    Wvalor9 = rstEspecificaciones!valor9
                    Wvalor10 = rstEspecificaciones!valor10
                    
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
                            !valor2 = Wvalor2
                            !Valor3 = WValor3
                            !valor4 = Wvalor4
                            !valor5 = Wvalor5
                            !valor6 = Wvalor6
                            !valor7 = Wvalor7
                            !valor8 = Wvalor8
                            !valor9 = Wvalor9
                            !valor10 = Wvalor10
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
    
    rstEspecificaciones.Close
    
    End If

    'moviienmtos varios de laboratiorio
        
    spmovlab = "ListamovlabTotal"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        
    coderr = 0
    With rstMovlab
            .MoveFirst
            Do
                If rstMovlab!WDate = WFectraspaso Then
                
                    WCodigo = rstMovlab!Codigo
                    WRenglon = rstMovlab!Renglon
                    WFecha = rstMovlab!Fecha
                    WFechaord = rstMovlab!Fechaord
                    WTipo = rstMovlab!Tipo
                    WArticulo = rstMovlab!Articulo
                    WTerminado = rstMovlab!Terminado
                    WCantidad = rstMovlab!Cantidad
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WClave = rstMovlab!Clave
                    WLote = rstMovlab!Lote
                    
                    With rstWMovlab
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
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
    
    rstMovlab.Close
    
    End If
    

    Rem PRODUCTOS TERMINADOS
        
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        
    coderr = 0
    With rstTerminado
            .MoveFirst
            Do
                If rstTerminado!WDate = WFectraspaso Then
                
                    WCodigo = rstTerminado!Codigo
                    WDescripcion = rstTerminado!Descripcion
                    WLinea = rstTerminado!Linea
                    WUnidad = rstTerminado!Unidad
                    WInicial = rstTerminado!Inicial
                    WEntradas = rstTerminado!Entradas
                    WSalidas = rstTerminado!Salidas
                    WMinimo = rstTerminado!Minimo
                    WDeposito = rstTerminado!Deposito
                    WPedido = rstTerminado!Pedido
                    WEnvase1 = rstTerminado!Envase1
                    WEnvase2 = rstTerminado!Envase2
                    WEnvase3 = rstTerminado!Envase3
                    WEnvase4 = rstTerminado!Envase4
                    WEnvase5 = rstTerminado!Envase5
                    WEnvase6 = rstTerminado!Envase6
                    WProceso = rstTerminado!Proceso
                    WCosto = rstTerminado!Costo
                    WFactor = rstTerminado!Factor
                    WImpreadi = rstTerminado!Impreadi
                    WClase = rstTerminado!Clase
                    WIntervencion = rstTerminado!Intervencion
                    WNaciones = rstTerminado!Naciones
                    WEmbalaje = rstTerminado!Embalaje
                    WVersion = rstTerminado!Version
                    WFechaversion = rstTerminado!Fechaversion
                    WDife = rstTerminado!Dife
                    
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
                            !Impreadi = WImpreadi
                            !Clase = WClase
                            !Intervencion = WIntervencion
                            !Naciones = WNaciones
                            !Embalaje = WEmbalaje
                            !Version = WVersion
                            !Fechaversion = WFechaversion
                            !Dife = WDife
                            
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
                            !Impreadi = WImpreadi
                            !Clase = WClase
                            !Intervencion = WIntervencion
                            !Naciones = WNaciones
                            !Embalaje = WEmbalaje
                            !Version = WVersion
                            !Fechaversion = WFechaversion
                            !Dife = WDife
                            
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
    
    rstTerminado.Close
    
    End If
    
   Rem precios por cliente
        
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        
    coderr = 0
    With rstPrecios
            .MoveFirst
            Do
                If rstPrecios!WDate = WFectraspaso Then
                
                    WCliente = rstPrecios!Cliente
                    WTerminado = rstPrecios!Terminado
                    WPrecio = rstPrecios!Precio
                    WClave = rstPrecios!Clave
                    WDescripcion = rstPrecios!Descripcion
                    WFecha1 = rstPrecios!Fecha1
                    WFactura1 = rstPrecios!Factura1
                    WPrecio1 = rstPrecios!Precio1
                    WCantidad1 = rstPrecios!Cantidad1
                    WFecha2 = rstPrecios!Fecha2
                    WFactura2 = rstPrecios!Factura2
                    WPrecio2 = rstPrecios!Precio2
                    WCantidad2 = rstPrecios!Cantidad2
                    WFecha3 = rstPrecios!Fecha3
                    WFactura3 = rstPrecios!Factura3
                    WPrecio3 = rstPrecios!Precio3
                    WCantidad3 = rstPrecios!Cantidad3
                    WFecha4 = rstPrecios!Fecha4
                    WFactura4 = rstPrecios!Factura4
                    WPrecio4 = rstPrecios!Precio4
                    WCantidad4 = rstPrecios!Cantidad4
                    WFecha5 = rstPrecios!Fecha5
                    WFactura5 = rstPrecios!Factura5
                    WPrecio5 = rstPrecios!Precio5
                    WCantidad5 = rstPrecios!Cantidad5
                    
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
    
    rstPrecios.Close
    
    End If
    
    
   Rem Composicion
        
    spComposicion = "ListaComposicion"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then
        
    coderr = 0
    With rstComposicion
            .MoveFirst
            Do
                If rstComposicion!WDate = WFectraspaso Then
                                        
                    WTerminado = rstComposicion!Terminado
                    WRenglon = rstComposicion!Renglon
                    WTipo = rstComposicion!Tipo
                    WArticulo1 = rstComposicion!Articulo1
                    WArticulo2 = rstComposicion!Articulo2
                    WCantidad = rstComposicion!Cantidad
                    WClave = rstComposicion!Clave
                    
                    With rstWComposicion
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Terminado = WTerminado
                            !Renglon = WRenglon
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
                            !Renglon = WRenglon
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
    
    rstComposicion.Close
    
    End If
    
   Rem Articulo
        
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        
    coderr = 0
    With rstArticulo
            .MoveFirst
            Do
                If rstArticulo!WDate = WFectraspaso Then
                
                    WCodigo = rstArticulo!Codigo
                    WDescripcion = rstArticulo!Descripcion
                    WCosto1 = rstArticulo!Costo1
                    WCosto2 = rstArticulo!Costo2
                    WInicial = rstArticulo!Inicial
                    WEntradas = rstArticulo!Entradas
                    WSalidas = rstArticulo!Salidas
                    WMinimo = rstArticulo!Minimo
                    WLaboratorio = rstArticulo!Laboratorio
                    WUnidad = rstArticulo!Unidad
                    WPedido = rstArticulo!Pedido
                    WDeposito = rstArticulo!Deposito
                    WEnvase = rstArticulo!Envase
                    WRs = rstArticulo!Rs
                    WProveedor = rstArticulo!Proveedor
                    WFecha = rstArticulo!Fecha
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
    
    rstArticulo.Close
    
    End If
    
   Rem cuenta corriente
        
    A1 = "      "
    A2 = "ZZZZZZ"
   
    XParam = "'" + A1 + "','" _
            + A2 + "'"
    spCtacte = "ListaCtacteDesdeHasta " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        
    coderr = 0
    With rstCtacte
            .MoveFirst
            Do
                If rstCtacte!WDate = WFectraspaso Then
                
                    WTipo = rstCtacte!Tipo
                    WNumero = rstCtacte!Numero
                    WRenglon = rstCtacte!Renglon
                    WCliente = rstCtacte!Cliente
                    WFecha = rstCtacte!Fecha
                    WEstado = rstCtacte!Estado
                    Wvencimiento = rstCtacte!Vencimiento
                    WVencimiento1 = rstCtacte!Vencimiento1
                    WTotal = rstCtacte!Total
                    WTotalUs = rstCtacte!TotalUs
                    WSaldo = rstCtacte!Saldo
                    WSaldoUs = rstCtacte!Saldous
                    WOrdFecha = rstCtacte!OrdFecha
                    WOrdVencimiento = rstCtacte!OrdVencimiento
                    WOrdVencimiento1 = rstCtacte!OrdVencimiento1
                    WImpre = rstCtacte!Impre
                    WNeto = rstCtacte!Neto
                    WIva1 = rstCtacte!Iva1
                    WIva2 = rstCtacte!Iva2
                    WPedido = rstCtacte!Pedido
                    WRemito = rstCtacte!Remito
                    WOrden = rstCtacte!Orden
                    WParidad = rstCtacte!Paridad
                    WProvincia = rstCtacte!Provincia
                    WVendedor = rstCtacte!Vendedor
                    WRubro = rstCtacte!Rubro
                    WComprobante = rstCtacte!Comprobante
                    WAceptada = rstCtacte!Aceptada
                    WCosto = rstCtacte!Costo
                    WImporte1 = rstCtacte!Importe1
                    WImporte2 = rstCtacte!Importe2
                    WImporte3 = rstCtacte!Importe3
                    WImporte4 = rstCtacte!Importe4
                    WImporte5 = rstCtacte!Importe5
                    WImporte6 = rstCtacte!Importe6
                    WImporte7 = rstCtacte!Importe7
                    WClave = rstCtacte!Clave
                    
                    With rstWCtaCte
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !Saldous = WSaldoUs
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
                            !Renglon = WRenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !Saldous = WSaldoUs
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
    
    rstCtacte.Close
    
    End If
    
    
   Rem ESTADISTICA
        
    A1 = "00000000"
    A2 = "99999999"
   
   
    XParam = "'" + A1 + "','" _
                 + A2 + "'"
    spEstadistica = "ListaEstadisticaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

    coderr = 0
    With rstEstadistica
            .MoveFirst
            Do
                If rstEstadistica!WDate = WFectraspaso Then
                
                    WTipo = rstEstadistica!Tipo
                    WNumero = rstEstadistica!Numero
                    WRenglon = rstEstadistica!Renglon
                    WArticulo = rstEstadistica!Articulo
                    WCantidad = rstEstadistica!Cantidad
                    WPrecio = rstEstadistica!Precio
                    WPrecioUs = rstEstadistica!PrecioUs
                    WImporte = rstEstadistica!Importe
                    WimporteUs = rstEstadistica!ImporteUs
                    WCliente = rstEstadistica!Cliente
                    WParidad = rstEstadistica!Paridad
                    WVendedor = rstEstadistica!Vendedor
                    WRubro = rstEstadistica!Rubro
                    WLinea = rstEstadistica!Linea
                    WCosto1 = rstEstadistica!Costo1
                    WCosto2 = rstEstadistica!Costo2
                    WCoeficiente = rstEstadistica!Coeficiente
                    WPedido = rstEstadistica!Pedido
                    WFecha = rstEstadistica!Fecha
                    WImporte1 = rstEstadistica!Importe1
                    WImporte2 = rstEstadistica!Importe2
                    WImporte3 = rstEstadistica!Importe3
                    WImporte4 = rstEstadistica!Importe4
                    WOrdFecha = rstEstadistica!OrdFecha
                    WWArticulo = rstEstadistica!WArticulo
                    WRemito = rstEstadistica!Remito
                    WClave = rstEstadistica!Clave
                    WLote1 = rstEstadistica!Lote1
                    WCanti1 = rstEstadistica!Canti1
                    WLote2 = rstEstadistica!Lote2
                    WCanti2 = rstEstadistica!Canti2
                    WLote3 = rstEstadistica!Lote3
                    WCanti3 = rstEstadistica!Canti3
                    WLote4 = rstEstadistica!Lote4
                    WCanti4 = rstEstadistica!Canti4
                    WLote5 = rstEstadistica!Lote5
                    WCanti5 = rstEstadistica!Canti5
                    
                    With rstWEstadistica
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WimporteUs
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
                            !Lote1 = WLote1
                            !Canti1 = WCanti1
                            !Lote2 = WLote2
                            !Canti2 = WCanti2
                            !Lote3 = WLote3
                            !Canti3 = WCanti3
                            !Lote4 = WLote4
                            !Canti4 = WCanti4
                            !Lote5 = WLote5
                            !Canti4 = WCanti5
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WimporteUs
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
                            !Lote1 = WLote1
                            !Canti1 = WCanti1
                            !Lote2 = WLote2
                            !Canti2 = WCanti2
                            !Lote3 = WLote3
                            !Canti3 = WCanti3
                            !Lote4 = WLote4
                            !Canti4 = WCanti4
                            !Lote5 = WLote5
                            !Canti4 = WCanti5
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
    
    rstEstadistica.Close
    
    End If
    
    
   Rem Desfripcion de comprobantes
        
   
    spDesccomp = "ConsultaDesccompTotal"
    Set rstDesccomp = db.OpenRecordset(spDesccomp, dbOpenSnapshot, dbSQLPassThrough)
    If rstDesccomp.RecordCount > 0 Then
        
    coderr = 0
    With rstDesccomp
            .MoveFirst
            Do
                If rstDesccomp!WDate = WFectraspaso Then
                
                    WTipo = rstDesccomp!Tipo
                    WNumero = rstDesccomp!Numero
                    WRenglon = rstDesccomp!Renglon
                    WDescripcion = rstDesccomp!Descripcion
                    WImporte = rstDesccomp!Importe
                    WClave = rstDesccomp!Clave
                    
                    With rstWDescComp
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Descripcion = WDescripcion
                            !Importe = WImporte
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
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
    
    rstDesccomp.Close
    
    End If
    
   Rem INFORME DE RECEPCION
        
    spInforme = "ListaInformeTotal"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        
    coderr = 0
    With rstInforme
            .MoveFirst
            Do
                If rstInforme!WDate = WFectraspaso Then
                
                    WInforme = rstInforme!Informe
                    WRenglon = rstInforme!Renglon
                    WFecha = rstInforme!Fecha
                    WProveedor = rstInforme!Proveedor
                    WRemito = rstInforme!Remito
                    WOrden = rstInforme!Orden
                    WArticulo = rstInforme!Articulo
                    WCantidad = rstInforme!Cantidad
                    WResta = rstInforme!Resta
                    WClave = rstInforme!Clave
                    WFechaord = rstInforme!Fechaord
                
                    With rstWInforme
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Informe = WInforme
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !Fechaord = WFechaord
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Informe = WInforme
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Remito = WRemito
                            !Orden = WOrden
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Resta = WResta
                            !Clave = WClave
                            !Fechaord = WFechaord
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
    
    rstInforme.Close
    
    End If
    
    
    
   Rem LAUDO DE LIBERACION
        
    spLaudo = "ListaLaudoTotal"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        
    coderr = 0
    With rstLaudo
            .MoveFirst
            Do
                If rstLaudo!WDate = WFectraspaso Then
                
                    WLaudo = rstLaudo!Laudo
                    WRenglon = rstLaudo!Renglon
                    WFecha = rstLaudo!Fecha
                    WOrden = rstLaudo!Orden
                    WArticulo = rstLaudo!Articulo
                    WLiberada = rstLaudo!Liberada
                    WDevuelta = rstLaudo!Devuelta
                    WLiberada = rstLaudo!Liberada
                    WLote = rstLaudo!Lote
                    WRechazo = rstLaudo!Rechazo
                    WActualiza = rstLaudo!Actualiza
                    WMarca = rstLaudo!Marca
                    WInforme = rstLaudo!Informe
                    WClave = rstLaudo!Clave
                    WSaldo = rstLaudo!Saldo
                
                    With rstWLaudo
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Laudo = WLaudo
                            !Renglon = WRenglon
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
                            !Saldo = WSaldo
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Laudo = WLaudo
                            !Renglon = WRenglon
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
                            !Saldo = WSaldo
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
    
    rstLaudo.Close
    
    End If
    
    
    
   Rem MOVIIMENTOS VARIOS
        
   
    spMovvar = "ListaMovvarTotal"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        
    coderr = 0
    With rstMovvar
            .MoveFirst
            Do
                If rstMovvar!WDate = WFectraspaso Then
                
                    WCodigo = rstMovvar!Codigo
                    WRenglon = rstMovvar!Renglon
                    WFecha = rstMovvar!Fecha
                    WFechaord = rstMovvar!Fechaord
                    WTipo = rstMovvar!Tipo
                    WArticulo = rstMovvar!Articulo
                    WTerminado = rstMovvar!Terminado
                    WCantidad = rstMovvar!Cantidad
                    WMovi = rstMovvar!Movi
                    WTipomov = rstMovvar!Tipomov
                    WObservaciones = rstMovvar!Observaciones
                    WClave = rstMovvar!Clave
                    WLote = rstMovvar!Lote
                
                    With rstWMovvar
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Lote = WLote
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
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
    
    rstMovvar.Close
    
    End If
    
    
    
   Rem MOVIIMENTOS VARIOS
   
    spMovguia = "ListaMovguiaTotal"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        
    coderr = 0
    With rstMovguia
            .MoveFirst
            Do
                If rstMovguia!WDate = WFectraspaso Then
                
                    WClave = rstMovguia!Clave
                    WTipomov = rstMovguia!Tipomov
                    WCodigo = rstMovguia!Codigo
                    WRenglon = rstMovguia!Renglon
                    WFecha = rstMovguia!Fecha
                    WTipo = rstMovguia!Tipo
                    WArticulo = rstMovguia!Articulo
                    WTerminado = rstMovguia!Terminado
                    WCantidad = rstMovguia!Cantidad
                    WFechaord = rstMovguia!Fechaord
                    WMovi = rstMovguia!Movi
                    WObservaciones = rstMovguia!Observaciones
                    WDestino = rstMovguia!Destino
                    WLote = rstMovguia!Lote
                    WSaldo = rstMovguia!Saldo
                    WPartida = rstMovguia!Partida

                    With rstWMovguia
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            !Destino = WDestino
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Partida = WPartida
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Movi = WMovi
                            !Tipomov = WTipomov
                            !Observaciones = WObservaciones
                            !Clave = WClave
                            !Destino = WDestino
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Partida = WPartida
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
    
    rstMovguia.Close
    
    End If
    
    
    
    
    
    
    

   Rem HOJA DE PRODUCCION
        
        
    spHoja = "ListaHojaTotal"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        
        
    coderr = 0
    With rstHoja
            .MoveFirst
            Do
                If rstHoja!WDate = WFectraspaso Then
                
                    WHoja = rstHoja!Hoja
                    WRenglon = rstHoja!Renglon
                    WFecha = rstHoja!Fecha
                    WProducto = rstHoja!Producto
                    WTeorico = rstHoja!Teorico
                    WReal = rstHoja!Real
                    WFechaing = rstHoja!fechaIng
                    WFechaingord = rstHoja!Fechaingord
                    WTipo = rstHoja!Tipo
                    WArticulo = rstHoja!Articulo
                    WTerminado = rstHoja!Terminado
                    WCantidad = rstHoja!Cantidad
                    WLote = rstHoja!Lote
                    WClave = rstHoja!Clave
                    WSaldo = rstHoja!Saldo
                    WLote1 = rstHoja!Lote1
                    WCanti1 = rstHoja!Canti1
                    WLote2 = rstHoja!Lote2
                    WCanti2 = rstHoja!Canti2
                    WLote3 = rstHoja!Lote3
                    WCanti3 = rstHoja!Canti3
                
                    With rstWHoja
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Hoja = WHoja
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WFechaing
                            !Fechaingord = WWfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Lote1 = WLote1
                            !Canti1 = WCanti1
                            !Lote2 = WLote2
                            !Canti2 = WCanti2
                            !Lote3 = WLote3
                            !Canti3 = WCanti3
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Hoja = WHoja
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Producto = WProducto
                            !Teorico = WTeorico
                            !Real = WReal
                            !fechaIng = WFechaing
                            !Fechaingord = WWfechaIngord
                            !Tipo = WTipo
                            !Articulo = WArticulo
                            !Terminado = WTerminado
                            !Cantidad = WCantidad
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Lote1 = WLote1
                            !Canti1 = WCanti1
                            !Lote2 = WLote2
                            !Canti2 = WCanti2
                            !Lote3 = WLote3
                            !Canti3 = WCanti3
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
    
    rstHoja.Close
    
    End If
    
    
    
    'Solicitud de ordenes de compra
        
    spSolic = "ListaSolicitudTotal"
    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolic.RecordCount > 0 Then
        
    coderr = 0
    With rstSolic
            .MoveFirst
            Do
                If rstSolic!WDate = WFectraspaso Then
                
                    WClave = rstSolic!Clave
                    WSolicitud = rstSolic!Solicitud
                    WRenglon = rstSolic!Renglon
                    WFecha = rstSolic!Fecha
                    WFechaord = rstSolic!Fechaord
                    WObservaciones = rstSolic!Observaciones
                    WArticulo = rstSolic!Articulo
                    WCantidad = rstSolic!Cantidad
                    WEntrega = rstSolic!Entrega
                    WOrdEntrega = rstSolic!OrdEntrega
                    WPlanta = rstSolic!Planta
                    WSolicitante = rstSolic!Solicitante
                    WDate = rstSolic!WDate
                    WMarca = rstSolic!Marca
                    WObser = rstSolic!Obser
                    WEntregado = rstSolic!Entregado
                    
                    With rstWSolic
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Clave = WClave
                            !Solicitud = WSolicitud
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Observaciones = WObservaciones
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Entrega = WEntrega
                            !OrdEntrega = WOrdEntrega
                            !Planta = WPlanta
                            !Solicitante = WSolicitante
                            !Date = WDate
                            !Marca = WMarca
                            !Obser = WObser
                            !Entregado = WEntregado
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Clave = WClave
                            !Solicitud = WSolicitud
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Fechaord = WFechaord
                            !Observaciones = WObservaciones
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Entrega = WEntrega
                            !OrdEntrega = WOrdEntrega
                            !Planta = WPlanta
                            !Solicitante = WSolicitante
                            !Date = WDate
                            !Marca = WMarca
                            !Obser = WObser
                            !Entregado = WEntregado
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
    
    rstSolic.Close
    
    End If

    
    
    
    
    
    
    
    
    
    
    With rstWSolic
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
    With rstWClientes
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
    
    DbsTraspa.Close
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub




