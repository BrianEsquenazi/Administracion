VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreOrden 
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
         Caption         =   "Pedidos a imprimir"
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
Attribute VB_Name = "PrgImpreOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim Vector(1000, 2) As String
Dim Datos(100, 10) As String
Dim WPedido As String
Dim WCliente As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WDespago As String
Dim WFecha As String
Dim WFecEntrega As String
Dim wversion As String
Dim WTipoped As String
Dim Lugar As Integer
Dim WEnvase(10) As String
Dim XEnvase(40, 6) As String
Dim Auxiliar(100, 4) As String
Dim WImpre(10) As String
Dim WArticulo As String
Dim WCantidad As Double
Dim WPartida As String
Dim Partida(100, 3) As String
Dim LugarPartida As String
Dim WSaldo As Double
Dim XLote(10, 10) As String
Dim WCantiProceso As String

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

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

    Erase XEnvase

    Lugar = 0

    spPedido = "ListaPedidoTotalListado4"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                        
                Entra = "S"
                            
                For XDa = 1 To Lugar
                    If Vector(Lugar, 1) = rstPedido!Pedido Then
                        Entra = "N"
                        Exit For
                    End If
                Next XDa
                                
                If Entra = "S" Then
                    Lugar = Lugar + 1
                    Vector(Lugar, 1) = rstPedido!Pedido
                    Vector(Lugar, 2) = "1"
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstPedido.Close
    
    End If
    
    spPedido = "ListaPedidoTotalListado2"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Rem If rstPedido!Renglon = 1 Then
                
                    If rstPedido!Autorizo = "X" Then
                    
                        If rstPedido!Impresion1 = "N" Then
                        
                            Entra = "S"
                            
                            For XDa = 1 To Lugar
                                If Vector(Lugar, 1) = rstPedido!Pedido Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next XDa
                                
                            If Entra = "S" Then
                                Lugar = Lugar + 1
                                Vector(Lugar, 1) = rstPedido!Pedido
                                Vector(Lugar, 2) = "2"
                            End If
                            
                        End If
                        
                    End If
                    
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstPedido.Close
    
    End If
    
    

    If Lugar > 0 Then
        PrgImpreped.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImpreped.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Rem Open "dada.txt" For Output As #1
    Open "lpt1" For Output As #1

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        Select Case Val(WTipoPedido)
            Case 1
                Call Proceso_Click
                Call Impresion
                
                WMarca = "X"
                XParam = "'" + WPedido + "','" _
                        + WMarca + "'"
                                           
                spPedido = "ModificaPedidoImpresion " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
                WMarca = "N"
                XParam = "'" + WPedido + "','" _
                        + WMarca + "'"
                                           
                spPedido = "ModificaPedidoImpresion1 " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                WMarca = "3"
                XParam = "'" + WPedido + "','" _
                             + WMarca + "'"
                                           
                spPedido = "ModificaPedidoProceso1 " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                Call ProcesoII_click
                Call ImpresionII
                WMarca = "X"
                                
                XParam = "'" + WPedido + "','" _
                             + WMarca + "'"
                
                spPedido = "ModificaPedidoImpresion1 " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                           
        End Select
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreped.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
        Rem WObservaciones = rstCliente!Observaciones
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + WPago + "'"
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
        Print #1, wversion;
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
                WProceso2 = Val(Datos(WCounter, 6))
                WCantiProceso = Datos(WCounter, 7)
                Call Ceros(WCantiProceso, 6)
                
                Select Case WProceso2
                    Case 1
                        WPartida = "Stock"
                    Case 2
                        WPartida = "Pda:" + WCantiProceso
                    Case 3
                        WPartida = "Transfer."
                    Case Else
                        WPartida = ""
                End Select
                    
                If WCantidad <> 0 Then
                    
                    Print #1, Tab(1); "|";
                    Print #1, Tab(2); WArticulo;
                    Print #1, Tab(16); "|";
                    Print #1, Tab(17); Left$(WDescripcion, 28);
                    Print #1, Tab(47); "|";
                    Print #1, Tab(48); WPartida;
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

End Sub


Private Sub Proceso_Click()

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
                        wversion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                        WTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                        WObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                        WObservaciones = Left$(WObservaciones + Space$(100), 100)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        Datos(Renglon, 6) = IIf(IsNull(rstPedido!proceso2), "", rstPedido!proceso2)
                        Datos(Renglon, 7) = IIf(IsNull(rstPedido!cantiproceso), "", rstPedido!cantiproceso)
                
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
    
    For da = 1 To WRenglon
    
        Cliente = Auxiliar(da, 1)
        Terminado = Auxiliar(da, 2)
        Tipopro = Auxiliar(da, 3)
        Articulo = Auxiliar(da, 4)
        
        Renglon = Renglon + 1
        
        If Left$(Terminado, 2) <> "DY" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Datos(Renglon, 1) = rstPrecios!Descripcion
                Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
                rstPrecios.Close
            End If
                Else
            spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + Articulo + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                Datos(Renglon, 3) = Pusing("###,###.##", rstPreciosMp!Precio)
                rstPreciosMp.Close
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Datos(Renglon, 1) = rstArticulo!Descripcion
                rstArticulo.Close
            End If
        End If
    Next da

End Sub

Private Sub ImpresionII()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
        Rem WObservaciones = rstCliente!Observaciones
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + WPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDespago = rstPago!Nombre
        End If
    End If
    
    Print #1, Tab(1); String$(79, "-")
        
    Print #1, Tab(1); "| SURFACTAN S.A.";
    Print #1, Tab(80); "|"
        
    Print #1, Tab(1); "|";
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Pedido";
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
    Print #1, Tab(58); "|";
    Print #1, Tab(59); "Cantidad";
    Print #1, Tab(67); "|";
    Print #1, Tab(68); "Partida";
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
                Print #1, Tab(67); "|";
                    
                XLote1 = XLote(WLugar, 1)
                XLote2 = XLote(WLugar, 3)
                XLote3 = XLote(WLugar, 5)
                XLote4 = XLote(WLugar, 7)
                XLote5 = XLote(WLugar, 9)
                XCantiLote1 = XLote(WLugar, 2)
                XCantiLote2 = XLote(WLugar, 4)
                XCantiLote3 = XLote(WLugar, 6)
                XCantiLote4 = XLote(WLugar, 8)
                XCantiLote5 = XLote(WLugar, 10)
                                
                For Cicla = 1 To 10 Step 2
                    If Val(XLote(WCounter, Cicla)) <> 0 Then
                        Select Case Cicla
                            Case 1
                                Print #1, Tab(68); Alinea("######", XLote(WCounter, Cicla));
                                Print #1, Tab(75); Alinea("#####", XLote(WCounter, Cicla + 1));
                                Print #1, Tab(80); "|"
    
                            Case Else
                                Print #1, Tab(1); "|";
                                Print #1, Tab(16); "|";
                                Rem Print #1, Tab(47); "|";
                                Print #1, Tab(58); "|";
                                Print #1, Tab(67); "|";
                                Print #1, Tab(68); Alinea("######", XLote(WCounter, Cicla));
                                Print #1, Tab(75); Alinea("#####", XLote(WCounter, Cicla + 1));
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
        Rem Print #1, Tab(47); "|";
        Print #1, Tab(58); "|";
        Print #1, Tab(67); "|";
        Print #1, Tab(80); "|"
    Next WDa
                
    Print #1, Tab(1); String$(79, "-")
    Print #1, Tab(1); "|";
    Print #1, Tab(10); "|";
    Print #1, Tab(18); "|";
    Print #1, Tab(26); "|";
    Print #1, Tab(34); "|";
    Print #1, Tab(42); "|";
    Print #1, Tab(50); "|";
    Print #1, Tab(58); "|";
    Print #1, Tab(66); "|";
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
    Print #1, Tab(1); ""
    
    Print #1, Chr$(12)

End Sub


Private Sub ProcesoII_click()

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
                        wversion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                        WTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                        WObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                        WObservaciones = Left$(WObservaciones + Space$(100), 100)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        
                        XLote(Renglon, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        XLote(Renglon, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        XLote(Renglon, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        XLote(Renglon, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        XLote(Renglon, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        XLote(Renglon, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        XLote(Renglon, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        XLote(Renglon, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        XLote(Renglon, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        XLote(Renglon, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                
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
    
    For da = 1 To WRenglon
    
        Cliente = Auxiliar(da, 1)
        Terminado = Auxiliar(da, 2)
        Tipopro = Auxiliar(da, 3)
        Articulo = Auxiliar(da, 4)
        
        Renglon = Renglon + 1
        
        If Left$(Terminado, 2) <> "DY" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Datos(Renglon, 1) = rstPrecios!Descripcion
                Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
                rstPrecios.Close
            End If
                Else
            spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + Articulo + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                Datos(Renglon, 3) = Pusing("###,###.##", rstPreciosMp!Precio)
                rstPreciosMp.Close
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Datos(Renglon, 1) = rstArticulo!Descripcion
                rstArticulo.Close
            End If
        End If
    Next da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Busca_Partida()

    Erase Partida
    LugarPartida = 0
    WPartida = ""
    
    XArti = Left$(WArticulo, 3) + Right$(WArticulo, 7)
    
    XParam = "'" + XArti + "','" _
                 + XArti + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                        Else
                    If rstLaudo!Articulo = XArti Then
                        WLote = rstLaudo!laudo
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        Call Redondeo(WSaldo)
                        WAno = Right$(!Fecha, 4)
                        WMes = Mid$(!Fecha, 4, 2)
                        WDia = Left$(!Fecha, 2)
                        WFecha = WAno + WMes + WDia
                        If WSaldo <> 0 Then
                            LugarPartida = LugarPartida + 1
                            Partida(LugarPartida, 1) = Str$(WLote)
                            Partida(LugarPartida, 2) = Str$(WSaldo)
                            Partida(LugarPartida, 3) = WFecha
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
        rstLaudo.Close
    End If
    
    
    XParam = "'" + XArti + "','" _
                + XArti + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
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
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = XArti Then
                        WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                        WAno = Right$(!Fecha, 4)
                        WMes = Mid$(!Fecha, 4, 2)
                        WDia = Left$(!Fecha, 2)
                        WFecha = WAno + WMes + WDia
                        If WSaldo <> 0 Then
                            LugarPartida = LugarPartida + 1
                            Partida(LugarPartida, 1) = Str$(WLaudo)
                            Partida(LugarPartida, 2) = Str$(WSaldo)
                            Partida(LugarPartida, 3) = WFecha
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
    
    For CicloPartida = 1 To LugarPartida
        For dada = CicloPartida + 1 To LugarPartida
            If Partida(CicloPartida, 3) > Partida(dada, 3) Then
                Auxi1 = Partida(CicloPartida, 1)
                Auxi2 = Partida(CicloPartida, 2)
                Auxi3 = Partida(CicloPartida, 3)
                
                Partida(CicloPartida, 1) = Partida(dada, 1)
                Partida(CicloPartida, 2) = Partida(dada, 2)
                Partida(CicloPartida, 3) = Partida(dada, 3)
                
                Partida(dada, 1) = Auxi1
                Partida(dada, 2) = Auxi2
                Partida(dada, 3) = Auxi3
            End If
        Next dada
    Next CicloPartida
    
    For CicloPartida = 1 To LugarPartida
        If Val(Partida(CicloPartida, 2)) >= WCantidad Then
            WPartida = Partida(CicloPartida, 1)
            Exit For
        End If
    Next CicloPartida
    
End Sub


