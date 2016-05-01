VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreOrdSii 
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
Attribute VB_Name = "PrgImpreOrdSii"
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

    WEmpresa = "0003"
    txtOdbc = "Empresa03"
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
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstOrden.Close
    
    End If
    
    If Lugar > 0 Then
        PrgImpreOrdSii.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImpreOrdSii.Refresh
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
                     + WMarca + "'"
                                           
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
    
    Rem Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreOrdSii.Hide
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
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, Orden.Corpeta, " + _
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

End Sub


Private Sub ImpresionAnterior()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", WEmpresa
        If .NoMatch = False Then
            Impretit = !Nombre
                Else
            Impretit = ""
        End If
    End With
    
    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WDesProveedor = RstProveedor!Nombre
        RstProveedor.Close
            Else
        DesProveedor = ""
    End If

    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, "--------------------------------------------------------------------------------"
        
    Print #1, Tab(1); "|";
    Print #1, Impretit;
    Print #1, Tab(80); "|"
        
    Print #1, Tab(1); "|";
    Print #1, Tab(60); "Remito :..........";
    Print #1, Tab(80); "|"
        
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Orden.....: ";
    Print #1, Tab(20); Alinea("######", WOrden);
    If WTipoOrden = 1 Then
        Print #1, Tab(30); "(IMPORTACION)";
    End If
    Print #1, Tab(50); "Fecha : "; WFecha;
    Print #1, Tab(80); "|"
        
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Proveedor...:"; Tab(20); WProveedor;
    Print #1, Tab(35); Left$(WDesProveedor, 20);
    Print #1, Tab(60); "Informe :.........";
    Print #1, Tab(80); "|"
        
    Print #1, Tab(1); "|";
        
    If Val(WCarpeta) <> 0 Then
        Print #1, Tab(5); "Carpeta.....:"; Tab(20); WCarpeta;
        Print #1, Tab(80); "|"
            Else
        Print #1, Tab(1); "|";
        Print #1, Tab(80); "|"
    End If
        
    Print #1, "--------------------------------------------------------------------------------"
    Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
    Print #1, "--------------------------------------------------------------------------------"

    XCantidad = 0
    Cantidad = 0
    Valor = 0
        
    For DA = 1 To 100
    
        WArticulo = Datos(DA, 1)
        WCantidad = Datos(DA, 3)
        WPrecio = Datos(DA, 4)
        WFecha1 = Datos(DA, 5)
        WFecha2 = Datos(DA, 6)
        WCondicion = Datos(DA, 7)
        
        If Left$(WArticulo, 2) <> "" And Left$(WArticulo, 2) <> Space$(2) Then
                
            XProveedor = WProveedor
            Call Ceros(XProveedor, 11)
            ClaveMarcas = WArticulo + XProveedor
            spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
            Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
            If rstMarcas.RecordCount > 0 Then
                WDescripcionMarcas = rstMarcas!Descripcion
                rstMarcas.Close
                    Else
                WDescripcionMarcas = ""
            End If
                    
            WUbicacion = ""
            WDescripcion = ""
                
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WUbicacion = rstArticulo!Deposito
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                
            XCantidad = XCantidad + 2

            Print #1, Tab(1); "|"; WArticulo;
            Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
            Print #1, Tab(41); "|"; WCantidad;
            Print #1, Tab(50); "|"; WFecha1;
            Print #1, Tab(61); "|"; WFecha2;
            Print #1, Tab(72); "|";
            Print #1, Tab(80); "|"
                        
            Print #1, Tab(1); "|";
            Print #1, Tab(12); "|"; WUbicacion;
            Print #1, Tab(50); "|";
            Print #1, Tab(61); "|";
            Print #1, Tab(72); "|";
            Print #1, Tab(80); "|"

        End If
        
    Next DA

    For Ciclo = XCantidad To 15
        Print #1, "|          |                            |        |          |          |       |"
    Next Ciclo

    Print #1, "--------------------------------------------------------------------------------"
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""

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

Private Sub Form_Activate()
    OPEN_FILE_Empresa
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
            DA = 0
            With rstEtiqueta
                .Index = "Codigo"
                .Seek ">=", DA
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
            
                DA = 0
                With rstEtiqueta
                    .Index = "Codigo"
                    .Seek ">=", DA
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
                    For DA = 1 To ZCantidad
                        .Index = "Codigo"
                        .AddNew
                    
                        ZLote = ""
                    
                        ZDa = Int((DA - 1) / 2)
                    
                        !Codigo = DA
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
                        If DA = ZCantidad And ZMarca = 1 Then
                            !Bruto = 1
                        End If
                        !Neto = ZDa
                        !Observaciones = "CUARENTENA"
                        .Update
                    Next DA
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
                
            
                DA = 0
                With rstEtiqueta
                    .Index = "Codigo"
                    .Seek ">=", DA
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




