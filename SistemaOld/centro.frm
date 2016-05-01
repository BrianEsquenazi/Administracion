VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCentro 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Pedidos Ingresados"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11850
   Begin VB.CommandButton Impre 
      Caption         =   "ReImpresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   4000
      Cols            =   7
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCentro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim WEnvase(10) As String
Dim XEnvase(40, 6) As String
Dim Datos(100, 10) As String
Dim Auxiliar(100, 4) As String
Dim WImpre(10) As String
Dim WPedido As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgCentro.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1200
    Muestra.ColWidth(2) = 1400
    Muestra.ColWidth(3) = 1400
    Muestra.ColWidth(4) = 4100
    Muestra.ColWidth(5) = 1400
    Muestra.ColWidth(6) = 1400
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    WPosi1 = 1
    WPosi2 = 1
    
End Sub

Private Sub Impre_Click()
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
        WPedido = Muestra.TextMatrix(Ciclo, 1)
        Call Lee_Click
    Next Ciclo
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If

    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Renglon = 0
    WSaldo = 0
    
    Pasa = 0
    Pedido = ""
    Fecha = "  /  /    "
    Cliente = ""
    Razon = ""
    FEntrega = "  /  /    "
    Tipo = 0
    Importe = 0
    Estado = ""
    
    Rem spPedido = "ListaPedidoCentro "
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    
    WPara1 = "N"
    
    Sql1 = "Select Pedido, Fecha, Cliente, FecEntrega, TipoPed, Autorizo, Impresion, Cantidad, Facturado, Precio, Impresion3, Terminado, Proceso1, TipoPedido"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.Proceso1 = 1"
    Sql4 = " and Pedido.Impresion = " + "'" + WPara1 + "'"
    If Val(XEmpresa) <> 1 Then
        Sql5 = " and Pedido.TipoPedido <> 5"
            Else
        Sql5 = " and Pedido.TipoPedido = 5"
    End If
    Sql6 = " Order by Pedido"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                If Pasa = 0 Then
                    Corte = rstPedido!Pedido
                    Fecha = rstPedido!Fecha
                    Cliente = rstPedido!Cliente
                    FEntrega = rstPedido!FecEntrega
                    Tipo = rstPedido!Tipoped
                    Importe = 0
                    Estado = rstPedido!Autorizo
                    Impresa = rstPedido!Impresion
                    Pasa = 1
                End If
                    
                If Corte <> rstPedido!Pedido Then
                    
                    Renglon = Renglon + 1
            
                    Muestra.Row = Renglon
                    
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(Corte))
                        
                    Muestra.Col = 2
                    Muestra.Text = Fecha
                
                    Muestra.Col = 3
                    Muestra.Text = Cliente
                        
                    Muestra.Col = 4
                    Muestra.Text = ""
                        
                    Muestra.Col = 5
                    Muestra.Text = FEntrega
                        
                    Select Case Tipo
                        Case 0
                            Muestra.Col = 6
                            Muestra.Text = "Normal"
                        Case 1
                            Muestra.Col = 6
                            Muestra.Text = "A Fecha"
                        Case 2
                            Muestra.Col = 6
                            Muestra.Text = "Fecha LImite"
                        Case 3
                            Muestra.Col = 6
                            Muestra.Text = "Urgente"
                        Case 4
                            Muestra.Col = 6
                            Muestra.Text = "Retiro Cliente"
                        Case 5
                            Muestra.Col = 6
                            Muestra.Text = "Muestra"
                        Case Else
                            Muestra.Col = 6
                            Muestra.Text = ""
                    End Select
                                                
                    Corte = rstPedido!Pedido
                    Fecha = rstPedido!Fecha
                    Cliente = rstPedido!Cliente
                    FEntrega = rstPedido!FecEntrega
                    Tipo = rstPedido!Tipoped
                    Importe = 0
                    Estado = rstPedido!Autorizo
                    Impresa = rstPedido!Impresion
                    Pasa = 1
                        
                End If
                    
                Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(Corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = FEntrega
                        
        Muestra.Col = 6
        Muestra.Text = Str$(Tipo)
        
        Select Case Tipo
            Case 0
                Muestra.Col = 6
                Muestra.Text = "Normal"
            Case 1
                Muestra.Col = 6
                Muestra.Text = "A Fecha"
            Case 2
                Muestra.Col = 6
                Muestra.Text = "Fecha LImite"
            Case 3
                Muestra.Col = 6
                Muestra.Text = "Urgente"
            Case 4
                Muestra.Col = 6
                Muestra.Text = "Retira Cliente"
            Case 5
                Muestra.Col = 6
                Muestra.Text = "Muestra"
            Case Else
                Muestra.Col = 6
                Muestra.Text = ""
        End Select
                        
    End If
    
    rstPedido.Close
    
    End If
    
    
    For dada = 1 To Renglon
    
        Muestra.Row = dada
                        
        Muestra.Col = 3
        WCliente = Muestra.Text
    
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.Col = 4
            Muestra.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next dada
    
    Call Conecta_Empresa
    
    TotalPedidos = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
End Sub

Private Sub Muestra_DblClick()

    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row

    Muestra.Col = 1
    WXPed = Muestra.Text
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        PrgPedCentro.Show
            Else
        PrgPedCentroPelli.Show
    End If
            
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    Call Proceso_Click
    Muestra.TopRow = WPosi1
    Muestra.Row = WPosi2
End Sub

Private Sub Lee_Click()

    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If

    Erase XEnvase
    Erase Datos
    Erase Auxiliar
    
    Renglon = 0
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    WCliente = rstPedido!Cliente
                    WFecha = rstPedido!Fecha
                    WFecEntrega = rstPedido!FecEntrega
                    WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                    WTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                    WObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                    WObservaciones = Left$(WObservaciones + Space$(100), 100)
                    ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                
                    Renglon = Renglon + 1
                        
                    Datos(Renglon, 0) = rstPedido!Terminado
                    Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                    Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                    Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                
                    XEnvase(Renglon, 1) = rstPedido!Envase1
                    XEnvase(Renglon, 2) = rstPedido!Canti1
                    XEnvase(Renglon, 3) = rstPedido!Envase2
                    XEnvase(Renglon, 4) = rstPedido!Canti2
                    XEnvase(Renglon, 5) = rstPedido!Envase3
                    XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                    WRenglon = WRenglon + 1
                    
                    Auxiliar(WRenglon, 1) = rstPedido!Cliente
                    Auxiliar(WRenglon, 2) = rstPedido!Terminado
                    Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                    Auxiliar(WRenglon, 4) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)

                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    Renglon = 0
    
    For Da = 1 To WRenglon
        Cliente = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Tipopro = Auxiliar(Da, 3)
        Articulo = Auxiliar(Da, 4)
        
        Renglon = Renglon + 1
        
        spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Datos(Renglon, 1) = rstPrecios!Descripcion
            Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
            rstPrecios.Close
        End If
    Next Da

    Rem Open "dada.txt" For Output As #1
    Open "lpt2" For Output As #1
    
    WEnvase(1) = 20
    WEnvase(2) = 21
    WEnvase(3) = 22
    WEnvase(4) = 23
    WEnvase(5) = 24
    WEnvase(6) = 25
    WEnvase(7) = 26
    WEnvase(8) = 30
    WEnvase(9) = 28

    For Cicla = 1 To 9
        spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
            rstEnvase.Close
                    Else
            WImpre(Cicla) = ""
        End If
    Next Cicla
    
    spCliente = "ConsultaClienteRazon " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        
        Rem WDirentrega = rstCliente!DirEntrega
        WDirentrega = ""
        
        ZDirEntrega(1) = rstCliente!DirEntrega
        ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WDirentrega = ZDirEntrega(ZLugarDirEntrega)
        
        Rem WObservaciones = rstCliente!Observaciones
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + Str$(WPago) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDespago = rstPago!Nombre
        End If
    End If

    For XX = 1 To 1

        Print #1, Tab(1); String$(79, "-")
        
        Print #1, Tab(1); "| SURFACTAN S.A.";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WPedido;
        Print #1, " / ";
        Print #1, WVersion;
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
        Print #1, Tab(5); "Fecha Ent.";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WFecEntrega;
        Select Case WTipoped
            Case 0
                Print #1, " (Normal)";
            Case 1
                Print #1, " (A fecha)";
            Case 2
                Print #1, " (Fecha Limite)";
            Case 3
                Print #1, " (Urgente)";
            Case 4
                Print #1, " (Retira Cliente)";
            Case 5
                Print #1, " (Muestra)";
            Case Else
        End Select
            
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "C.Pago";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDespago;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Entrega";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDirentrega;
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
        Print #1, Tab(47); "|";
        Print #1, Tab(48); "Partida";
        Print #1, Tab(58); "|";
        Print #1, Tab(59); "Cantidad";
        Print #1, Tab(67); "|";
        Print #1, Tab(68); "Envase";
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
                WPartida = ""
                    
                If WCantidad <> 0 Then
                    
                    Print #1, Tab(1); "|";
                    Print #1, Tab(2); WArticulo;
                    Print #1, Tab(16); "|";
                    Print #1, Tab(17); Left$(WDescripcion, 28);
                    Print #1, Tab(47); "|";
                    Print #1, Tab(50); WPartida;
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Alinea("###,###", Str$(WCantidad));
                    Print #1, Tab(67); "|";
                                
                    For Cicla = 1 To 6 Step 2
                        If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                            Select Case Cicla
                                Case 1
                                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvase.RecordCount > 0 Then
                                        WAbre = rstEnvase!Abreviatura
                                            Else
                                        WAbre = ""
                                    End If
                                
                                    Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                    Print #1, Tab(80); "|"
                                    
                                Case Else
                                    Print #1, Tab(1); "|";
                                    Print #1, Tab(16); "|";
                                    Print #1, Tab(47); "|";
                                    Print #1, Tab(58); "|";
                                    Print #1, Tab(67); "|";
                                                
                                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvase.RecordCount > 0 Then
                                        WAbre = rstEnvase!Abreviatura
                                            Else
                                        WAbre = ""
                                    End If
                                            
                                    Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                    Print #1, Tab(80); "|"
                                    XLinea = XLinea + 1
                                                
                            End Select
                        End If
                    Next Cicla
                    XLinea = XLinea + 1
                    
                End If
                    
            End If
            
        Next WCounter
        
        For WDa = XLinea To 10
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(47); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|"; WImpre(1);
        Print #1, Tab(10); "|"; WImpre(2);
        Print #1, Tab(18); "|"; WImpre(3);
        Print #1, Tab(26); "|"; WImpre(4);
        Print #1, Tab(34); "|"; WImpre(5);
        Print #1, Tab(42); "|"; WImpre(6);
        Print #1, Tab(50); "|"; WImpre(7);
        Print #1, Tab(58); "|"; WImpre(8);
        Print #1, Tab(66); "|"; WImpre(9);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|(020)";
        Print #1, Tab(10); "|(021)";
        Print #1, Tab(18); "|(022)";
        Print #1, Tab(26); "|(023)";
        Print #1, Tab(34); "|(024)";
        Print #1, Tab(42); "|(025)";
        Print #1, Tab(50); "|(026)";
        Print #1, Tab(58); "|(030)";
        Print #1, Tab(66); "|(028)";
        Print #1, Tab(80); "|"
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""

    Next XX
    
    Print #1, Chr$(12)
    
    Close #1
    
    Call Conecta_Empresa

End Sub


