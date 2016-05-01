VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListPresu 
   Caption         =   "Listado de Proyeccion de Dy"
   ClientHeight    =   6600
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6600
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "DY-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "DY-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "listpresu.frx":0000
      Left            =   0
      List            =   "listpresu.frx":0007
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListPresu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(10000, 4) As String
Dim WFechaCierre As String
Dim WOrdFechaCierre As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private xLote(100, 7) As String
Dim MesActual As String
Dim MesAnterior As String
Dim AnoAnterior As String
Dim AnoActual As String

Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    With rstStockDy
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    Rem
    Rem recalcula saldo de orden de compra
    Rem
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spArticulo = "ModificaArticuloPedido0DesdeHasta" + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Sql1 = "Select Clave, Orden, Articulo, Cantidad, FechaOrd"
    Sql2 = " FROM Orden"
    Sql3 = " Where Orden.Articulo >= " + "'" + Desde.Text + "'"
    Sql4 = " and Orden.Articulo <= " + "'" + Hasta.Text + "'"
    Sql5 = " and Orden.FechaOrd > " + "'" + "20040101" + "'"
    spOrden = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstOrden!Clave
                    Vector(Renglon, 2) = rstOrden!Orden
                    Vector(Renglon, 3) = rstOrden!Articulo
                    Vector(Renglon, 4) = rstOrden!Cantidad
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WClave = Vector(Ciclo, 1)
        WOrden = Vector(Ciclo, 2)
        WArticulo = Vector(Ciclo, 3)
        WCantidad = Val(Vector(Ciclo, 4))
        WResta = 0
        
        XParam = "'" + WOrden + "','" _
                 + WArticulo + "'"
        spInforme = "ListaInformeOrdenArticulo" + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            With rstInforme
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                        WResta = WResta + rstInforme!Resta
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstInforme.Close
        End If
        
        WDife = WCantidad - WResta
        XParam = "'" + WArticulo + "','" _
                     + Str$(WDife) + "'"
        spArticulo = "ModificaArticuloPedido " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    Rem
    Rem recalcula saldo de informe de recepcion
    Rem
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spArticulo = "ModificaArticuloLaboratorio0DesdeHasta" + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Sql1 = "Select Clave, Informe, Articulo, Cantidad, FechaOrd"
    Sql2 = " FROM Informe"
    Sql3 = " Where Informe.Articulo >= " + "'" + Desde.Text + "'"
    Sql4 = " and Informe.Articulo <= " + "'" + Hasta.Text + "'"
    Sql5 = " and Informe.FechaOrd > " + "'" + "20040101" + "'"
    spInforme = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstInforme!Clave
                    Vector(Renglon, 2) = rstInforme!Informe
                    Vector(Renglon, 3) = rstInforme!Articulo
                    Vector(Renglon, 4) = rstInforme!Cantidad
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If

    For Ciclo = 1 To Renglon
    
        WClave = Vector(Ciclo, 1)
        WInforme = Vector(Ciclo, 2)
        WArticulo = Vector(Ciclo, 3)
        WCantidad = Val(Vector(Ciclo, 4))
        WResta = 0
        
        XParam = "'" + WInforme + "','" _
                 + WArticulo + "'"
        spLaudo = "ListaLaudoInforme " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                        WLiberada = rstLaudo!Liberada
                        WDevuelta = rstLaudo!devuelta
                        WSuma = WLiberada + WDevuelta
                        WLiberadaAnt = IIf(IsNull(rstLaudo!liberadaant), "0", rstLaudo!liberadaant)
                        WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        WSumaAnt = WLiberadaAnt + WDevueltaAnt
                        If WSumaAnt <> 0 Then
                            WResta = WResta + WSumaAnt
                                Else
                            WResta = WResta + WSuma
                        End If
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstLaudo.Close
        End If
        
        WDife = WCantidad - WResta
        XParam = "'" + WArticulo + "','" _
                     + Str$(WDife) + "'"
        spArticulo = "ModificaArticuloLaboratorio " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    Rem
    Rem calcula los pedidos pendientess
    Rem
    
    spArticulo = "ModificaArticuloVenta0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ModificaPedpenDy " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ListaPedidoPendDesdeHasta " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    EntraVector = "S"
                    For Ciclo = 1 To LugarVector
                        If Vector(Ciclo, 1) = rstPedido!Terminado Then
                            Vector(Ciclo, 2) = Str$(Val(Vector(Ciclo, 2)) + rstPedido!Importe)
                            EntraVector = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If EntraVector = "S" Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = rstPedido!Terminado
                        Vector(Renglon, 2) = Str$(rstPedido!Importe)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To Renglon
        WProducto = Vector(Ciclo, 1)
        WImporte = Vector(Ciclo, 2)
        WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
        XParam = "'" + WArticulo + "','" _
                     + WImporte + "','" _
                     + WDate + "'"
        spArticulo = "ModificaArticuloVenta " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    
    


    Rem
    Rem recalcula saldo del stock
    Rem

    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spArticulo = "ListaArticuloDesdeHasta " + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    With rstArticulo
            .MoveFirst
            Do
                If .EOF = True Then
                    Exit Do
                End If
                Renglon = Renglon + 1
                Vector(Renglon, 1) = rstArticulo!Codigo
                Vector(Renglon, 2) = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
                Vector(Renglon, 3) = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstArticulo.Close
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WArticulo = Vector(Da, 1)
        XCodigo = Vector(Da, 1)
        WFechaCierre = Vector(Da, 2)
        WOrdFechaCierre = Vector(Da, 3)
        XDate = Date$
        
        Call calcula_datos
        
        XEntradas = Str$(WEntradas)
        XSalidas = Str$(WSalidas)
        
        XParam = "'" + XCodigo + "','" _
                + XEntradas + "','" _
                + XSalidas + "','" _
                + XDate + "'"
                                           
        spArticulo = "ModificaArticuloMovimientos " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da
    
    
    Rem
    Rem reproceso de calculo de datos para el listado
    Rem
    
        
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
    spArticulo = "ListaArticuloDesdeHasta" + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
            
        With rstArticulo
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    WArticulo = rstArticulo!Codigo
                    WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    WDescripcion = rstArticulo!Descripcion
                    WLaboratorio = rstArticulo!Laboratorio
                    WOrden = rstArticulo!Pedido
                    WPedido = rstArticulo!Venta
                    WMinimo = rstArticulo!Minimo
                
                    With rstStockDy
                        .Index = "Codigo"
                        .Seek "=", WArticulo
                        If .NoMatch = True Then
                            .AddNew
                            !Codigo = WArticulo
                            !Descripcion = WDescripcion
                            !Stock1 = 0
                            !Stock2 = 0
                            !Stock3 = 0
                            !Stock4 = 0
                            !Stock5 = 0
                            !Stock1 = !Stock1 + WStock
                            !Titulo1 = "Surfactan S.A."
                            !Titulo2 = "al " + Fecha.Text
                            !Laboratorio = WLaboratorio
                            !Orden = WOrden
                            !Pedido = WPedido
                            !Familia = Mid$(WArticulo, 4, 3)
                            !Minimo = WMinimo
                            .Update
                                Else
                            .Edit
                            !Stock1 = !Stock1 + WStock
                            !Laboratorio = WLaboratorio
                            !Orden = WOrden
                            !Pedido = WPedido
                            !Familia = Mid$(WArticulo, 4, 3)
                            !Minimo = WMinimo
                            .Update
                        End If
                    End With
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        End With
            
        rstArticulo.Close
    End If
    
    
    
    With rstStockDy
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WArticulo = !Codigo
                
                MesActual = Mid$(Fecha.Text, 4, 2)
                AnoActual = Mid$(Fecha.Text, 7, 4)
                MesAnterior = Str$(Val(MesActual) - 1)
                AnoAnterior = Mid$(Fecha.Text, 7, 4)
                If Val(MesAnterior) = 0 Then
                    MesAnterior = "12"
                    AnoAnterior = Str$(Val(AnoAnterior) - 1)
                    Call Ceros(AnoAnterior, 4)
                End If
                Call Ceros(MesAnterior, 2)
                
                WImpo1 = 0
                WImpo2 = 0
                WImpo3 = 0
                
                Sql1 = "Select *"
                Sql2 = " FROM AltaProyec"
                Sql3 = " Where AltaProyec.Ano = " + "'" + AnoAnterior + "'"
                Sql4 = " and AltaProyec.Codigo = " + "'" + WArticulo + "'"
                spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                If rstAltaProyec.RecordCount > 0 Then
                    Select Case Val(MesAnterior)
                        Case 1
                            WImpo1 = rstAltaProyec!Mes1
                        Case 2
                            WImpo1 = rstAltaProyec!Mes2
                        Case 3
                            WImpo1 = rstAltaProyec!Mes3
                        Case 4
                            WImpo1 = rstAltaProyec!Mes4
                        Case 5
                            WImpo1 = rstAltaProyec!Mes5
                        Case 6
                            WImpo1 = rstAltaProyec!Mes6
                        Case 7
                            WImpo1 = rstAltaProyec!Mes7
                        Case 8
                            WImpo1 = rstAltaProyec!Mes8
                        Case 9
                            WImpo1 = rstAltaProyec!Mes9
                        Case 10
                            WImpo1 = rstAltaProyec!Mes10
                        Case 11
                            WImpo1 = rstAltaProyec!Mes11
                        Case 12
                            WImpo1 = rstAltaProyec!Mes12
                        Case Else
                    End Select
                    rstAltaProyec.Close
                End If
                
                Sql1 = "Select *"
                Sql2 = " FROM AltaProyec"
                Sql3 = " Where AltaProyec.Ano = " + "'" + AnoActual + "'"
                Sql4 = " and AltaProyec.Codigo = " + "'" + WArticulo + "'"
                spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                If rstAltaProyec.RecordCount > 0 Then
                    Select Case Val(MesActual)
                        Case 1
                            WImpo2 = rstAltaProyec!Mes1
                        Case 2
                            WImpo2 = rstAltaProyec!Mes2
                        Case 3
                            WImpo2 = rstAltaProyec!Mes3
                        Case 4
                            WImpo2 = rstAltaProyec!Mes4
                        Case 5
                            WImpo2 = rstAltaProyec!Mes5
                        Case 6
                            WImpo2 = rstAltaProyec!Mes6
                        Case 7
                            WImpo2 = rstAltaProyec!Mes7
                        Case 8
                            WImpo2 = rstAltaProyec!Mes8
                        Case 9
                            WImpo2 = rstAltaProyec!Mes9
                        Case 10
                            WImpo2 = rstAltaProyec!Mes10
                        Case 11
                            WImpo2 = rstAltaProyec!Mes11
                        Case 12
                            WImpo2 = rstAltaProyec!Mes12
                        Case Else
                    End Select
                    rstAltaProyec.Close
                End If
                
                WDesdefecha = AnoAnterior + MesAnterior + "01"
                WHastafecha = AnoAnterior + MesAnterior + "31"
                
                Sql1 = "Select Clave, Orden, Articulo, Cantidad, FechaOrd"
                Sql2 = " FROM Orden"
                Sql3 = " Where Orden.Articulo = " + "'" + WArticulo + "'"
                Sql4 = " and Orden.FechaOrd >= " + "'" + WDesdefecha + "'"
                Sql5 = " and Orden.FechaOrd <= " + "'" + WHastafecha + "'"
                spOrden = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                Renglon = Renglon + 1
                                WImpo3 = WImpo3 + rstOrden!Cantidad
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstOrden.Close
                End If
                
                !Stock2 = WImpo1
                !Stock3 = WImpo2
                !Stock4 = WImpo3
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    
    

    
    
    
    
    
    
    
    

    Listado.WindowTitle = "Listado de Proyeccion de DY"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WListProyec.rpt"
    
    Listado.Action = 1
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstStockDy
        .Close
    End With
    DbsEmpresa.Close
    
    Fecha.SetFocus
    PrgListproyec.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SelStart = 3
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SelStart = 3
        Desde.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_StockDy
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListproyec.Caption = "Listado de Proyeccion de Dy :  " + !Nombre
        End If
    End With
    
    Desde.Text = "DY-   -   "
    Desde.SelStart = 3
    Hasta.Text = "DY-   -   "
    Hasta.SelStart = 3
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstArticulo!Codigo, 2) = "DY" Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstArticulo
            
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                    Desde.Text = rstArticulo!Codigo
                    Hasta.Text = rstArticulo!Codigo
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Left$(rstArticulo!Codigo, 2) = "DY" Then
                    Da = Len(rstArticulo!Descripcion) - WEspacios
                    For aaa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, aaa, WEspacios) Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aaa
                End If
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub

Private Sub calcula_datos()


    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoRepro" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                        Else
                    If rstLaudo!Articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!Liberada
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spHoja = "ListaHojaRepro" + XParam
    Rem spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                        Else
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                        XX = rstHoja!Clave
                        WSalidas = WSalidas + rstHoja!Cantidad
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                If rstHoja!Articulo > WArticulo Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstHoja.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarRepro1" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovvar!Marca = "X" Then
                        Else
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaRepro1" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                        Else
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + rstMovguia!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovguia!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spMovlab = "ListaMovlabRepro1" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovlab!Marca = "X" Then
                        Else
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovlab!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
    End If
    
    
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spEstadistica = "ListaEstadisticaReproDy " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstEstadistica!Marca = "X" Then
                        Else
                    WTipo = rstEstadistica!Tipo
                    xLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    xLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    xLote(2, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote2)
                    xLote(2, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti2)
                    xLote(3, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote3)
                    xLote(3, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti3)
                    xLote(4, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote4)
                    xLote(4, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti4)
                    xLote(5, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote5)
                    xLote(5, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti5)
                        
                    For Da = 1 To 5
                        WLote = xLote(Da, 1)
                        WCantidad = xLote(Da, 2)
                        If Val(WCantidad) <> 0 Then
                            If WTipo = 2 Then
                                WEntradas = WEntradas + Abs(Val(WCantidad))
                                    Else
                                WSalidas = WSalidas + WCantidad
                            End If
                        End If
                    Next Da
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
    End If
    
End Sub





