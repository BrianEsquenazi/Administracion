VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreOrdSi 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Ordenes a imprimir"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImpreOrdSi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim XParam As String
Dim Vector(1000) As String
Dim Datos(100, 10) As String
Dim Lugar As Integer
Dim WOrden As String
Dim WTipoOrden As Integer
Dim WFecha As String
Dim WProveedor As String
Dim WDesProveedor As String
Dim WCarpeta As String
Dim XProveedor As String

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Lugar = 0

    XParam = "'" + "N" + "'"
    spOrden = "ListaOrdenImpresion " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
    With rstOrden
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If rstOrden!Orden < 900000 Then
                        
                Entra = "S"
                            
                For XDa = 1 To Lugar
                    If Vector(Lugar) = rstOrden!Orden Then
                        Entra = "N"
                        Exit For
                    End If
                Next XDa
                                
                If Entra = "S" Then
                    Lugar = Lugar + 1
                    Vector(Lugar) = rstOrden!Orden
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstOrden.Close
    
    End If
    
    Rem Lugar = Lugar + 1
    Rem Vector(Lugar) = 8330
    Rem Lugar = Lugar + 1
    Rem Vector(Lugar) = 8331
    Rem Lugar = Lugar + 1
    Rem Vector(Lugar) = 8332
    Rem Lugar = Lugar + 1
    Rem Vector(Lugar) = 8333
    
    If Lugar > 0 Then
        PrgImpreOrdSi.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImpreOrdSi.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Rem Open "dada.txt" For Output As #1
    Rem Open "lpt1" For Output As #1
    
    m$ = "Coloque  el papel para la impresion de las ordenes de compra"
    a% = MsgBox(m$, 0, "Impresion de Ordenes de Compra")

    For WWCicla = 1 To Lugar
    
        WOrden = Vector(WWCicla)
        
        Call Proceso_Click
        Call Impresion
                
        Rem WMarca = "S"
        Rem XParam = "'" + WOrden + "','" _
        rem              + WMarca + "'"
                                           
        Rem spOrden = "ModificaOrdenImpresion " + XParam
        Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    
    m$ = "Coloque el papel para la impresion de  las etiquetas internas"
    a% = MsgBox(m$, 0, "Impresion de Ordenes de Compra")

    For WWCicla = 1 To Lugar
    
        WOrden = Vector(WWCicla)
        
        Call Impre_Etiquetas
                
        WMarca = "S"
        XParam = "'" + WOrden + "','" _
                     + WMarca + "'"
                                           
        spOrden = "ModificaOrdenImpresion " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreOrdSi.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + WOrden + " to " + WOrden
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "OrdenImpreInterno.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, " + _
                            "Articulo.Descripcion, Proveedor.Nombre, Proveedor.CategoriaI " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + WOrden + " AND " + _
                            "Orden.Orden <= " + WOrden + " "
                            
    Listado.Connect = Connect()
    Listado.Action = 1

    Rem With rstEmpresa
    Rem     .Index = "Empresa"
    Rem     .Seek "=", WEmpresa
    Rem     If .NoMatch = False Then
    Rem         Impretit = !Nombre
    Rem             Else
    Rem         Impretit = ""
    Rem     End If
    Rem End With
    Rem
    Rem spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
    Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem If RstProveedor.RecordCount > 0 Then
    Rem     WDesProveedor = RstProveedor!Nombre
    Rem     RstProveedor.Close
    Rem         Else
    Rem     DesProveedor = ""
    Rem End If
    Rem
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, "--------------------------------------------------------------------------------"
    Rem
    Rem Print #1, Tab(1); "|";
    Rem Print #1, Impretit;
    Rem Print #1, Tab(80); "|"
    Rem
    Rem Print #1, Tab(1); "|";
    Rem Print #1, Tab(60); "Remito :..........";
    Rem Print #1, Tab(80); "|"
    Rem
    Rem Print #1, Tab(1); "|";
    Rem Print #1, Tab(5); "Orden.....: ";
    Rem Print #1, Tab(20); Alinea("######", WOrden);
    Rem If WTipoOrden = 1 Then
    Rem     Print #1, Tab(30); "(IMPORTACION)";
    Rem End If
    Rem Print #1, Tab(50); "Fecha : "; WFecha;
    Rem Print #1, Tab(80); "|"
    Rem
    Rem Print #1, Tab(1); "|";
    Rem Print #1, Tab(5); "Proveedor...:"; Tab(20); WProveedor;
    Rem Print #1, Tab(35); Left$(WDesProveedor, 20);
    Rem Print #1, Tab(60); "Informe :.........";
    Rem Print #1, Tab(80); "|"
    Rem
    Rem If Val(WCarpeta) <> 0 Then
    Rem     Print #1, Tab(1); "|";
    Rem     Print #1, Tab(5); "Carpeta.....:"; Tab(20); WCarpeta;
    Rem     Print #1, Tab(80); "|"
    Rem         Else
    Rem     Print #1, Tab(1); "|";
    Rem     Print #1, Tab(80); "|"
    Rem End If
    Rem
    Rem Print #1, "--------------------------------------------------------------------------------"
    Rem Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
    Rem Print #1, "--------------------------------------------------------------------------------"
    Rem
    Rem XCantidad = 0
    Rem Cantidad = 0
    Rem Valor = 0
    Rem
    Rem For Da = 1 To 100
    Rem
    Rem     WArticulo = Datos(Da, 1)
    Rem     WCantidad = Datos(Da, 3)
    Rem     WPrecio = Datos(Da, 4)
    Rem     WFecha1 = Datos(Da, 5)
    Rem     WFecha2 = Datos(Da, 6)
    Rem     WCondicion = Datos(Da, 7)
    Rem
    Rem     If Left$(WArticulo, 2) <> "" And Left$(WArticulo, 2) <> Space$(2) Then
    Rem
    Rem         XProveedor = WProveedor
    Rem         Call Ceros(XProveedor, 11)
    Rem         ClaveMarcas = WArticulo + XProveedor
    Rem         spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
    Rem         Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    Rem         If rstMarcas.RecordCount > 0 Then
    Rem             WDescripcionMarcas = rstMarcas!Descripcion
    Rem             rstMarcas.Close
    Rem                 Else
    Rem             WDescripcionMarcas = ""
    Rem         End If
    Rem
    Rem         WUbicacion = ""
    Rem         WDescripcion = ""
    Rem
    Rem         spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Rem         Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem         If rstArticulo.RecordCount > 0 Then
    Rem             WUbicacion = rstArticulo!Deposito
    Rem             WDescripcion = rstArticulo!Descripcion
    Rem             rstArticulo.Close
    Rem         End If
    Rem
    Rem         XCantidad = XCantidad + 2
    Rem
    Rem         Print #1, Tab(1); "|"; WArticulo;
    Rem         Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
    Rem         Print #1, Tab(41); "|"; WCantidad;
    Rem         Print #1, Tab(50); "|"; WFecha1;
    Rem         Print #1, Tab(61); "|"; WFecha2;
    Rem         Print #1, Tab(72); "|";
    Rem         Print #1, Tab(80); "|"
    Rem
    Rem         Print #1, Tab(1); "|";
    Rem         Print #1, Tab(12); "|"; WUbicacion;
    Rem         Print #1, Tab(50); "|";
    Rem         Print #1, Tab(61); "|";
    Rem         Print #1, Tab(72); "|";
    Rem         Print #1, Tab(80); "|"
    Rem
    Rem     End If
    Rem
    Rem Next Da
    Rem
    Rem For Ciclo = XCantidad To 15
    Rem     Print #1, "|          |                            |        |          |          |       |"
    Rem Next Ciclo
    Rem
    Rem Print #1, "--------------------------------------------------------------------------------"
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""
    Rem Print #1, ""

End Sub

Private Sub Proceso_Click()

    Renglon = 0
    Erase Datos
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
                
                WTipoOrden = rstOrden!Tipo
                WFecha = rstOrden!Fecha
                WProveedor = rstOrden!Proveedor
                WCarpeta = rstOrden!Carpeta
            
                Datos(Renglon, 1) = rstOrden!Articulo
                Datos(Renglon, 3) = Pusing("###,###.##", rstOrden!Cantidad)
                Datos(Renglon, 4) = Pusing("###,###.##", rstOrden!Precio)
                Datos(Renglon, 5) = rstOrden!Fecha1
                Datos(Renglon, 6) = rstOrden!fecha2
                Datos(Renglon, 7) = rstOrden!Condicion
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub


Private Sub Impre_Etiquetas()

    On Error GoTo WError
    
    OPEN_FILE_Etiqueta
    
    Renglon = 0
    Erase Datos
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
                Datos(Renglon, 1) = rstOrden!Articulo
                Datos(Renglon, 2) = IIf(IsNull(rstOrden!Bultos), "", rstOrden!Bultos)
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    
    For Ciclo = 1 To Renglon
    
        WProducto = Datos(Ciclo, 1)
        WCantiEti = Val(Datos(Ciclo, 2))
        
        If WCantiEti > 0 Then
    
            Salida = "N"
            Da = 0
            With rstEtiqueta
                .Index = "Codigo"
                .Seek ">=", Da
                If .NoMatch = False Then
                    Do
                        m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        Salida = "S"
                        Exit Do
                    Loop
                End If
            End With
            
            If Salida <> "S" Then
            
                WClase = ""
                WIntervencion = ""
                WNaciones = ""
                WEmbalaje = ""
                WDesProducto = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WProducto + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDesProducto = IIf(IsNull(rstArticulo!Descripcion), "", rstArticulo!Descripcion)
                    WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                    WIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                    WNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
                    WEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
                    rstArticulo.Close
                End If
            
                Da = 0
                With rstEtiqueta
                    .Index = "Codigo"
                    .Seek ">=", Da
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
                
                ZCantidad = Int(WCantiEti / 2)
                ZMarca = 0
                If ZCantidad * 2 <> WCantiEti Then
                    ZCantidad = ZCantidad + 1
                    ZMarca = 1
                End If
                
                With rstEtiqueta
                    For Da = 1 To ZCantidad
                        .Index = "Codigo"
                        .AddNew
                    
                        ZLote = ""
                    
                        ZDa = Int((Da - 1) / 2)
                    
                        !Codigo = Da
                        !Terminado = WProducto
                        !Lote = 0
                        !Cliente = ""
                        !Cantidad = 0
                        !Nombre = ""
                        !Impre1 = ""
                        !Razon = "Orden : " + WOrden
                        !DirEntrega = ""
                        !Clase = WClase
                        !Intervencion = WIntervencion
                        !Naciones = WNaciones
                        !Embalaje = WEmbalaje
                        !Bruto = 0
                        If Da = ZCantidad And ZMarca = 1 Then
                            !Bruto = 1
                        End If
                        !Neto = ZDa
                        !Observaciones = "CUARENTENA"
                        .Update
                    Next Da
                End With
    
                Listado.WindowTitle = "Emision de Etiquetas"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
            
                Select Case Mid$(WClase, 1, 1)
                    Case "3"
                        Listado.ReportFileName = "WEtiVerde3.rpt"
                    Case "5"
                        Listado.ReportFileName = "WEtiVerde5.rpt"
                    Case "6"
                        Listado.ReportFileName = "WEtiVerde6.rpt"
                    Case "8"
                        Listado.ReportFileName = "WEtiVerde8.rpt"
                    Case "9"
                        Listado.ReportFileName = "WEtiVerde9.rpt"
                    Case Else
                        Listado.ReportFileName = "WEtiVerde.rpt"
                End Select
                    
                Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
                Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
                Rem Listado.Connect = Connect()
        
                Listado.GroupSelectionFormula = ""
                Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
        
                Listado.Destination = 1
                Rem Listado.Destination = 0
                Listado.PrinterCopies = 1
                Listado.Action = 1
                
            
                Da = 0
                With rstEtiqueta
                    .Index = "Codigo"
                    .Seek ">=", Da
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
                
            End If
            
        End If
    
    Next Ciclo
    
    Exit Sub

WError:

    Resume Next

End Sub

