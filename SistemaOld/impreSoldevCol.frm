VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreSoldevCol 
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
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Solicitud de Devoluciones a Imprimir"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         TabIndex        =   2
         Top             =   240
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
Attribute VB_Name = "PrgImpreSoldevCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedidoDevol As Recordset
Dim spPedidoDevol As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim Vector(1000) As String
Dim Datos(100, 10) As String
Dim WPedido As String
Dim WCliente As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WFecha As String
Dim WFecEntrega As String
Dim Lugar As Integer
Dim Auxiliar(100, 4) As String
Dim WImpre(10) As String
Dim WArticulo As String
Dim WCantidad As Double
Dim WCantiProceso As String
Dim WObserva As String
Dim ZVector(100, 10) As String

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Lugar = 0

    spPedidoDevol = "ListaPedidoDevolTotalListado"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
    With rstPedidoDevol
        .MoveFirst
        If .NoMatch = False Then
            Do
                        
                WTipoPedido = IIf(IsNull(rstPedidoDevol!tipopedido), "", rstPedidoDevol!tipopedido)
                        
                If WTipoPedido = "CO" Or WTipoPedido = "PG" Then
                        
                Entra = "S"
                            
                For XDa = 1 To Lugar
                    If Vector(Lugar) = rstPedidoDevol!Pedido Then
                        Entra = "N"
                        Exit For
                    End If
                Next XDa
                                
                If Entra = "S" Then
                    Lugar = Lugar + 1
                    Vector(Lugar) = rstPedidoDevol!Pedido
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstPedidoDevol.Close
    End If
    
    If Lugar > 0 Then
        PrgImpreSoldevCol.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImpreSoldevCol.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Rem Open "dada.txt" For Output As #1
    Rem Open "lpt1" For Output As #1

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla)
        Call Proceso_Click
        Rem Call Impresion
        Call ImpresionSql
        
        WMarca = "X"
        XParam = "'" + WPedido + "','" _
                     + WMarca + "'"
                                           
        spPedidoDevol = "ModificaPedidoDevolImpresion " + XParam
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreSoldevCol.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WDirentrega = rstCliente!DirEntrega
        rstCliente.Close
    End If

    For XX = 1 To 1

        Print #1, Tab(1); String$(79, "-")
        
        Print #1, Tab(1); "| SURFACTAN S.A.";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "| Solicitud de Devolucion de Mercaderia";
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Solicitud Nro.";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WPedido;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Cliente";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WCliente;
        Print #1, Tab(40); Left$(WRazon, 35);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WFecha;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Left$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Right$(WObservaciones, 50);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|";
        Print #1, Tab(2); "Producto";
        Print #1, Tab(16); "|";
        Print #1, Tab(17); "Descripcion";
        Print #1, Tab(58); "|";
        Print #1, Tab(59); "Cantidad";
        Print #1, Tab(80); "|"

        Print #1, Tab(1); String$(79, "-")
        
        XLinea = 0
        WCounter = 0
                    
        For WCounter = 1 To 40
        
            If Datos(WCounter, 0) <> "" Then
                    
                WArticulo = Datos(WCounter, 0)
                WDescripcion = Datos(WCounter, 1)
                WCantidad = Val(Datos(WCounter, 2))
                WPrecio = Val(Datos(WCounter, 3))
                    
                If WCantidad <> 0 Then
                    
                    Print #1, Tab(1); "|";
                    Print #1, Tab(2); WArticulo;
                    Print #1, Tab(16); "|";
                    Print #1, Tab(17); Left$(WDescripcion, 28);
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Alinea("###,###", Str$(WCantidad));
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                    
                End If
                    
            End If
            
        Next WCounter
        
        For WDa = XLinea To 10
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(80); "|"
        Next WDa
                
        Print #1, Tab(1); String$(79, "-")

    Next XX
    
    Print #1, Chr$(12)

End Sub

Private Sub ImpresionSql()

    On Error GoTo WError
    
    Erase ZVector
    Renglon = 0

    spPedidoDevol = "ListaPedidoDevol " + "'" + WPedido + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
            With rstPedidoDevol
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZZCliente = rstPedidoDevol!Cliente
                        ZZFecha = rstPedidoDevol!Fecha
                        ZZObservaciones = IIf(IsNull(rstPedidoDevol!Observaciones), "", rstPedidoDevol!Observaciones)
                        ZZObservaciones = Left$(WObservaciones + Space$(100), 100)
                
                        Renglon = Renglon + 1
                        
                        ZVector(Renglon, 1) = rstPedidoDevol!articulo
                        ZVector(Renglon, 3) = rstPedidoDevol!cantidad
                        ZVector(Renglon, 4) = rstPedidoDevol!partida
                        ZVector(Renglon, 5) = rstPedidoDevol!TipoPro
                        ZVector(Renglon, 6) = rstPedidoDevol!Terminado
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedidoDevol.Close
    End If
    
        
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    WObservaciones = Left$(RTrim(Observaciones.Text) + Space$(100), 100)
        
    XLinea = 0
    
    For a = 1 To Renglon
        
        WArticulo = ZVector(a, 1)
        WCantidad = ZVector(a, 3)
        WPartida = ZVector(a, 4)
        WTipoPro = ZVector(a, 5)
        WTerminado = ZVector(a, 6)
        WDescripcion = ""
        WImpreCodigo = ""
        
        If WTipoPro = "M" Then
        
            WImpreCodigo = WArticulo
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
                Else
                
            WImpreCodigo = WTerminado
            ZClave = ZZCliente + WTerminado
            
            spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                WDescripcion = rstPrecios!Descripcion
                rstPrecios.Close
            End If
                    
        End If
        
        
        WRenglon = a
        
        Auxi = WPedido
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = WPedido
        ZRenglon = Str$(WRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = ""
        ZCliente = ZZCliente
        ZFecha = ZZFecha
        ZFechaent = "  /  /    "
        ZTipoPedido = WPartida
        ZCondicion = ""
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = ""
        ZArticulo = WImpreCodigo
        ZDescripcion = WDescripcion
        ZPrecio = ""
        ZCantidad = Str$(WCantidad)
        ZEnvase = ""
       Rem by nan 17-01-13
       Rem ZNombre = ""
        ZEntrega = ""
        
        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZNombre = rstCliente!Razon
            ZEntrega = rstCliente!DirEntrega
            rstCliente.Close
        End If
        
        spImprePed = "INSERT INTO ImprePed (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                    
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
            
    Next a
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
                    
        Auxi = WPedido
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = WPedido
        ZRenglon = Str$(WRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = ""
        ZCliente = ZZCliente
       Rem by nan 17-01-13
        Rem ZNombre = ZZDesCliente
        ZFecha = WFecha
        ZFechaent = "  /  /    "
        ZTipoPedido = ""
        ZCondicion = ""
        ZEntrega = WDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = ""
        ZArticulo = ""
        ZDescripcion = ""
        ZPrecio = ""
        ZCantidad = ""
        ZEnvase = ""
                    
        spImprePed = "INSERT INTO ImprePed (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Empresa, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.TipoPedido, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Cantidad " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999"
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "ImprepedidevolSQL.rpt"
    Listado.Destination = 1
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
    Exit Sub
        
WError:
    Resume Next

End Sub



Private Sub Proceso_Click()

    Erase Datos
    Erase Auxiliar
    
    Renglon = 0
    WRenglon = 0

    spPedidoDevol = "ListaPedidoDevol " + "'" + WPedido + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedidoDevol.RecordCount > 0 Then
            With rstPedidoDevol
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WCliente = rstPedidoDevol!Cliente
                        WFecha = rstPedidoDevol!Fecha
                        WObservaciones = IIf(IsNull(rstPedidoDevol!Observaciones), "", rstPedidoDevol!Observaciones)
                        WObservaciones = Left$(WObservaciones + Space$(100), 100)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedidoDevol!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedidoDevol!cantidad - rstPedidoDevol!Facturado)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedidoDevol!articulo), "", rstPedidoDevol!articulo)
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedidoDevol!Cliente
                        Auxiliar(WRenglon, 2) = rstPedidoDevol!Terminado
                        Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedidoDevol!TipoPro), "", rstPedidoDevol!TipoPro)
                        Auxiliar(WRenglon, 4) = IIf(IsNull(rstPedidoDevol!articulo), "", rstPedidoDevol!articulo)

                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedidoDevol.Close
    End If
    
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Cliente = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        TipoPro = Auxiliar(Da, 3)
        articulo = Auxiliar(Da, 4)
        
        Renglon = Renglon + 1
        
        If Left$(Terminado, 2) <> "DY" And Left$(Terminado, 2) <> "DW" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Datos(Renglon, 1) = rstPrecios!Descripcion
                Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
                rstPrecios.Close
            End If
                Else
            spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + articulo + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                Datos(Renglon, 3) = Pusing("###,###.##", rstPreciosMp!Precio)
                rstPreciosMp.Close
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Datos(Renglon, 1) = rstArticulo!Descripcion
                rstArticulo.Close
            End If
        End If
    Next Da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

