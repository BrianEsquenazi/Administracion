VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImprePigmentos 
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
Attribute VB_Name = "PrgImprePigmentos"
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
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim Vector(1000, 3) As String
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

Dim WEspecif(100) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim AuxiliarII(100, 5) As String

Dim ZZDirEntrega(10) As String
Dim WWEspecif(100) As String
Dim xLote(100, 22) As String
Dim ImpreEnvase(10) As String

Dim ZClave As String
Dim ZTipo As String
Dim ZPedido As String
Dim ZRenglon As String
Dim ZEmpresa As String
Dim ZVersion As String
Dim ZCliente As String
Dim ZNombre As String
Dim ZFecha As String
Dim ZFechaent As String
Dim ZTipoPedido As String
Dim ZCondicion As String
Dim ZEntrega As String
Dim ZObservaciones1 As String
Dim ZObservaciones2 As String
Dim ZOrden As String
Dim ZArticulo As String
Dim ZDescripcion As String
Dim ZPrecio As String
Dim ZCantidad As String
Dim ZEnvase As String
Dim ZZLote As String
Dim ZZCantiLote As String



Private Sub Acepta_Click()

    Rem WEmpresa = "0009"
    Rem txtOdbc = "Empresa09"
    
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

    spPedido = "ListaPedidoPigmentos"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                        
                If rstPedido!Autorizo = "X" Then
                    Entra = "S"
                        
                    For XDa = 1 To Lugar
                        If Vector(Lugar, 1) = rstPedido!pedido Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XDa
                                
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstPedido!pedido
                        Vector(Lugar, 2) = "1"
                        Vector(Lugar, 3) = rstPedido!Tipoped
                    End If
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
    
    
    
    
    
    
    
    
    
    
    
    Pasa = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Muestra"
    Sql3 = " Where Autoriza = " + "'" + "S" + "'"
    Sql4 = " and Impresion = " + "'" + "X" + "'"
    Sql5 = " Order by Pedido, Codigo"
    spMuestra = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        With rstMuestra
            .MoveFirst
            If .NoMatch = False Then
                Do
        
                    XTipoPro = ""
    
                    WTerminado = Trim(rstMuestra!Producto)
                    WArticulo = Trim(rstMuestra!Articulo)
    
                    If WTerminado <> "" Then
                        XCodigo = Val(Mid$(WTerminado, 4, 5))
                        If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DW" Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 0 And XCodigo <= 999 Then
                                XTipoPro = "CO"
                                    Else
                                If XCodigo >= 11000 And XCodigo <= 11999 Then
                                    XTipoPro = "CO"
                                        Else
                                    If XCodigo >= 25000 And XCodigo <= 25999 Then
                                        XTipoPro = "FA"
                                            Else
                                        If XCodigo >= 2300 And XCodigo <= 2399 Then
                                            XTipoPro = "BI"
                                                Else
                                            XTipoPro = "PT"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
    
                    If WArticulo <> "" Then
                       If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Or Left$(WArticulo, 2) = "CO" Then
                            XTipoPro = "CO"
                                Else
                            XTipoPro = "PT"
                        End If
                    End If
    
                    If XTipoPro = "PG" Then
    
                    If Pasa = 0 Then
                        Pasa = 1
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!pedido
                        Vector(Lugar, 2) = "2"
                        Corte = rstMuestra!pedido
                    End If
                    If Corte <> rstMuestra!pedido Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!pedido
                        Vector(Lugar, 2) = "2"
                        Vector(Lugar, 3) = ""
                        Corte = rstMuestra!pedido
                    End If
    
                    End If
    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        rstMuestra.Close
    End If
    
    
    
    
    
    
    

    If Lugar > 0 Then
        PrgImprePigmentos.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImprePigmentos.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Rem Open "lpt1" For Output As #1
    Rem Open "dada.txt" For Output As #1

    For WWCicla = 1 To Lugar
        
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        WTipoped = Vector(WWCicla, 3)
        
        Select Case Val(WTipoPedido)
            Case 1
                Rem Call Proceso_Click
                Rem If Val(WTipoped) = 5 Then
                Rem     Call ImpresionIII
                Rem         Else
                Rem     Call Impresion
                Rem End If
                
                Call ImpresionSql
                
                WMarca = "S"
            
                XParam = "'" + WPedido + "','" _
                        + WMarca + "'"
                                           
                spPedido = "ModificaPedidoPigmentos " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 2
                Call ProcesoIII_click
                Call ImpresionIII
                WMarca = "S"
                        
                Sql1 = "UPDATE Muestra SET "
                Sql2 = " Impresion =  " + "'" + WMarca + "'"
                Sql3 = " Where Pedido = " + "'" + WPedido + "'"
                spMuestra = Sql1 + Sql2 + Sql3
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                           
        End Select
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImprePigmentos.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!razon
        WPago = rstCliente!Pago1
        
        Rem WDirentrega = rstCliente!DirEntrega
        WDirentrega = ""
        
        ZDirEntrega(1) = rstCliente!DirEntrega
        ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WDirentrega = ZDirEntrega(ZLugarDirEntrega)
        
        Erase WEspecif
        WEspecif(1) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
        WEspecif(2) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
        WEspecif(3) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
        WEspecif(4) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
        WEspecif(5) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
        WEspecif(6) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
        WEspecif(7) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
        WEspecif(8) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
        WEspecif(9) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
        WEspecif(10) = IIf(IsNull(rstCliente!Especif10), "", rstCliente!Especif10)
        For CicloEspecif = 1 To 10
            WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
        Next CicloEspecif
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
        
        Pasa = 0
        For CicloEspecif = 1 To 10
            If WEspecif(CicloEspecif) <> "" Then
                If Pasa = 0 Then
                    Print #1, Tab(1); "|Especificaciones : ";
                    Pasa = 1
                End If
                Print #1, Tab(25); WEspecif(CicloEspecif);
                Print #1, Tab(80); "|"
                XLinea = XLinea + 1
            End If
        Next CicloEspecif
        
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

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub


Private Sub ImpresionIII()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
        WVendedor = Str$(rstCliente!Vendedor)
        Erase WEspecif
        WEspecif(1) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
        WEspecif(2) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
        WEspecif(3) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
        WEspecif(4) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
        WEspecif(5) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
        WEspecif(6) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
        WEspecif(7) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
        WEspecif(8) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
        WEspecif(9) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
        WEspecif(10) = IIf(IsNull(rstCliente!Especif10), "", rstCliente!Especif10)
        For CicloEspecif = 1 To 10
            WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
        Next CicloEspecif
        Rem WObservaciones = rstCliente!Observaciones
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + WPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDespago = rstPago!Nombre
            rstPago.Close
        End If
        
        spVendedor = "ConsultaVendedor " + "'" + WVendedor + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            WDesVendedor = rstVendedor!Nombre
            rstVendedor.Close
        End If
        
    End If

    Print #1, Tab(1); String$(79, "-")
        
    Print #1, Tab(1); "| SURFACTAN S.A.";
    Print #1, Tab(80); "|"
            
    Print #1, Tab(1); "|";
    Print #1, Tab(80); "|"
    
    Print #1, Tab(1); "| M U E S T R A S      P A R A     C L I E N T E S";
    Print #1, Tab(80); "|"
            
    Print #1, Tab(1); "|";
    Print #1, Tab(80); "|"
            
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Pedido";
    Print #1, Tab(28); ":";
    Print #1, Tab(30); WPedido;
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Fecha Pedido";
    Print #1, Tab(28); ":";
    Print #1, Tab(30); WFecha;
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Cliente";
    Print #1, Tab(28); ":";
    Print #1, Tab(30); WCliente;
    Print #1, Tab(40); Left$(WRazon, 35);
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Vendedor";
    Print #1, Tab(28); ":";
    Print #1, Tab(30); WVendedor;
    Print #1, Tab(40); Left$(WDesVendedor, 35);
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); "|";
    Print #1, Tab(5); "Observaciones";
    Print #1, Tab(28); ":";
    Print #1, Tab(30); Left$(WObservaciones, 50);
    Print #1, Tab(80); "|"
                
    Print #1, Tab(1); String$(79, "-")
    Print #1, Tab(1); "|";
    Print #1, Tab(2); "Producto";
    Print #1, Tab(16); "|";
    Print #1, Tab(17); "Descripcion";
    Print #1, Tab(47); "|";
    Print #1, Tab(48); "Cantidad";
    Print #1, Tab(60); "|";
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
            If WCantidad <> 0 Then
                Print #1, Tab(1); "|";
                Print #1, Tab(2); WArticulo;
                Print #1, Tab(16); "|";
                Print #1, Tab(17); Left$(WDescripcion, 28);
                Print #1, Tab(47); "|";
                Print #1, Tab(48); Pusing("#####.##", Str$(WCantidad));
                Print #1, Tab(60); "|";
                Print #1, Tab(80); "|"
                XLinea = XLinea + 1
            End If
        End If
    Next WCounter
    
    Pasa = 0
    For CicloEspecif = 1 To 10
        If WEspecif(CicloEspecif) <> "" Then
            If Pasa = 0 Then
                Print #1, Tab(1); "|Especificaciones : ";
                Pasa = 1
            End If
            Print #1, Tab(25); WEspecif(CicloEspecif);
            Print #1, Tab(80); "|"
            XLinea = XLinea + 1
        End If
    Next CicloEspecif
    
    For WDa = XLinea To 8
        Print #1, Tab(1); "|";
        Print #1, Tab(16); "|";
        Print #1, Tab(47); "|";
        Print #1, Tab(60); "|";
        Print #1, Tab(80); "|"
    Next WDa
                
    Print #1, Tab(1); String$(79, "-")
    Print #1, Tab(1); "|Preparó:";
    Print #1, Tab(21); "|Etiquetó:";
    Print #1, Tab(41); "|Registró";
    Print #1, Tab(61); "|Retiró:";
    Print #1, Tab(80); "|"
    Print #1, Tab(1); String$(79, "-")
    
    Print #1, Chr$(12)

End Sub

Private Sub ProcesoIII_click()

    Erase Datos
    
    Renglon = 0

    Sql1 = "Select *"
    Sql2 = " FROM Muestra"
    Sql3 = " Where Muestra.Pedido = " + "'" + WPedido + "'"
    Sql4 = " Order by Muestra.Codigo"
    spMuestra = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        With rstMuestra
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WVendedor = Str$(rstMuestra!Vendedor)
                    WObservaciones = Trim(rstMuestra!Observaciones)
                    WFecha = rstMuestra!Fecha
                    WCliente = rstMuestra!Cliente
                    WRazon = Trim(rstMuestra!razon)
                    
                    WProducto = Trim(rstMuestra!Producto)
                    WArticulo = Trim(rstMuestra!Articulo)
                    WEnsayo = Trim(rstMuestra!Ensayo)
                    
                    Renglon = Renglon + 1
                        
                    If WProducto <> "" Then
                        Datos(Renglon, 0) = WProducto
                        Datos(Renglon, 4) = "PT"
                            Else
                        If WArticulo <> "" Then
                            Datos(Renglon, 0) = WArticulo
                            Datos(Renglon, 4) = "MP"
                                Else
                            Datos(Renglon, 0) = WEnsayo
                            Datos(Renglon, 4) = "EN"
                        End If
                    End If
                    Datos(Renglon, 1) = Trim(rstMuestra!descricliente)
                    Datos(Renglon, 2) = Pusing("###,###.##", rstMuestra!Cantidad)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMuestra.Close
    End If
        
    spVendedor = "ConsultaVendedor " + "'" + WVendedor + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        WDesVendedor = rstVendedor!Nombre
        rstVendedor.Close
    End If
    
End Sub

Private Sub ImpresionSql()

    Rem On Error GoTo WError
    
    Erase XEnvase
    Erase Datos
    Erase AuxiliarII
    
    Renglon = 0
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WWCliente = rstPedido!Cliente
                    WWFecha = rstPedido!Fecha
                    WWFecEntrega = rstPedido!FecEntrega
                    WWVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                    WWTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                    WWObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                    WWObservaciones = Left$(WWObservaciones + Space$(100), 100)
                    ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                    ZZVia = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                    WWOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                    
                    Rem If rstPedido!Cantidad - rstPedido!Facturado > 0 Then
                    If rstPedido!Cantidad > 0 Then
            
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", Str$(rstPedido!Cantidad - rstPedido!Facturado))
                        Rem Datos(Renglon, 2) = Pusing("###,###.##", Str$(rstPedido!Cantidad))
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        Datos(Renglon, 6) = IIf(IsNull(rstPedido!proceso2), "", rstPedido!proceso2)
                        Datos(Renglon, 7) = IIf(IsNull(rstPedido!cantiproceso), "", rstPedido!cantiproceso)
                        Datos(Renglon, 8) = IIf(IsNull(rstPedido!observa), "", rstPedido!observa)
                        Datos(Renglon, 9) = IIf(IsNull(rstPedido!Especificaciones), "", rstPedido!Especificaciones)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        AuxiliarII(WRenglon, 1) = rstPedido!Cliente
                        AuxiliarII(WRenglon, 2) = rstPedido!Terminado
                        AuxiliarII(WRenglon, 3) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        AuxiliarII(WRenglon, 4) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                            AuxiliarII(WRenglon, 5) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
                        End If
                        
                    End If

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
    
        WWCliente = AuxiliarII(Da, 1)
        Cliente = AuxiliarII(Da, 1)
        Terminado = AuxiliarII(Da, 2)
        Tipopro = AuxiliarII(Da, 3)
        Articulo = AuxiliarII(Da, 4)
        ZZNombreComercial = AuxiliarII(Da, 5)
        
        Renglon = Renglon + 1
        
        If Left$(Terminado, 2) = "PT" Or Left$(Terminado, 2) = "YQ" Or Left$(Terminado, 2) = "YF" Then
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
            
            If ZZNombreComercial <> "" Then
                Datos(Renglon, 1) = ZZNombreComercial
                    Else
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Datos(Renglon, 1) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            End If
        End If
    Next Da
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImprePedIpPigmento"
    spImprePedIp = ZSql
    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    WObservaciones = Left$(WWObservaciones + Space$(100), 100)
    Select Case WWTipoped
        Case 0
            WTipoPedido = " (Normal)"
        Case 1
            WTipoPedido = " (A fecha)"
        Case 2
            WTipoPedido = " (Fecha Limite)"
        Case 3
            WTipoPedido = " (Urgente)"
        Case 4
            WTipoPedido = " (Retira Cliente)"
        Case 5
            WTipoPedido = " (Muestra)"
        Case Else
    End Select
    
    WVia = ""
    Select Case ZZVia
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
    
    
    spCliente = "ConsultaCliente " + "'" + Cliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        WWPago = rstCliente!Pago1
        WWDirentrega = ""
        WWDesCliente = rstCliente!razon
        
        ZZDirEntrega(1) = rstCliente!DirEntrega
        ZZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WWDirentrega = ZZDirEntrega(ZZLugarDirEntrega)
        
        rstCliente.Close
        
        spPago = "ConsultaPago " + "'" + Str$(WWPago) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WWDesPago = rstPago!Nombre
            rstPago.Close
        End If
        
    End If
    
    
    
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For A = 1 To 40
        
        WCounter = WCounter + 1
                
        If Datos(WCounter, 0) <> "" Then
                
            WArticulo = Datos(WCounter, O)
            WDescripcion = Datos(WCounter, 1)
            WCantidad = Val(Datos(WCounter, 2))
            WPrecio = Val(Datos(WCounter, 3))
            WObserva = Datos(WCounter, 8)
            WEspecificaciones = Datos(WCounter, 9)
                
            If WCantidad <> 0 Then
            
                Erase ImpreEnvase
                LugarEnvase = 0
            
                For Cicla = 1 To 6 Step 2
                    If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                        LugarEnvase = LugarEnvase + 1
                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnvase.RecordCount > 0 Then
                            WAbre = rstEnvase!Abreviatura
                            rstEnvase.Close
                                Else
                            WAbre = ""
                        End If
                        ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8)
                    End If
                Next Cicla
                
                WRenglon = WRenglon + 1
                
                Auxi = WPedido
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = WPedido
                ZRenglon = Str$(WRenglon)
                ZEmpresa = ""
                ZVersion = Str$(WWVersion)
                ZCliente = WWCliente
                ZNombre = WWDesCliente
                ZFecha = WWFecha
                ZFechaent = WWFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = WWDesPago
                ZEntrega = WWDirentrega
                
                ZObservaciones1 = Left$(WObservaciones, 50)
                ZObservaciones2 = Right$(WObservaciones, 50)
                ZOrden = WWOrdenCpa
                
                ZArticulo = WArticulo
                ZDescripcion = WDescripcion
                ZPrecio = Str$(WPrecio)
                ZCantidad = Str$(WCantidad)
                ZEnvase = ImpreEnvase(1)
                ZLugarLote = 1
                
                ZZLote = ""
                ZZCantiLote = ""
                Select Case ZLugarLote
                    Case 1
                        ZZLote = xLote(WCounter, 1)
                        ZZCantiLote = xLote(WCounter, 2)
                    Case 2
                        ZZLote = xLote(WCounter, 3)
                        ZZCantiLote = xLote(WCounter, 4)
                    Case 3
                        ZZLote = xLote(WCounter, 5)
                        ZZCantiLote = xLote(WCounter, 6)
                    Case 4
                        ZZLote = xLote(WCounter, 7)
                        ZZCantiLote = xLote(WCounter, 8)
                    Case 5
                        ZZLote = xLote(WCounter, 9)
                        ZZCantiLote = xLote(WCounter, 10)
                    Case Else
                End Select
                
                spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                            "Cantidad , Envase, Lote, CantiLote )" + _
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
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                For Ciclo = 2 To LugarEnvase
                
                    WRenglon = WRenglon + 1
                
                    Auxi = WPedido
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = WPedido
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = ""
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = WWDesCliente
                    ZFecha = WWFecha
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = ""
                    ZDescripcion = ""
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ImpreEnvase(Ciclo)
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    
                    
                    spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                            "Cantidad , Envase, Lote, CantiLote )" + _
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
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                    
                Next Ciclo
                
                If Trim(WEspecificaciones) <> "" And WEspecificaciones <> "0" Then
                
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
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = WWDesCliente
                    ZFecha = WWFecha
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = "Especif.:"
                    ZDescripcion = WEspecificaciones
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                            "Cantidad , Envase, Lote, CantiLote )" + _
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
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Trim(WObserva) <> "" Then
                
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
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = WWDesCliente
                    ZFecha = WWFecha
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = "Observ.:"
                    ZDescripcion = WObserva
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                            "Cantidad , Envase, Lote, CantiLote )" + _
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
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                For Ciclo = ZLugarLote + 1 To 5
                
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case Ciclo
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    If Val(ZZLote) <> 0 Or Val(ZZCantiLote) Then
                
                        WRenglon = WRenglon + 1
                    
                        Auxi = WPedido
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        ZClave = "1" + Auxi + Auxi1
                        ZTipo = "1"
                        ZPedido = WPedido
                        ZRenglon = Str$(WRenglon)
                        ZEmpresa = ""
                        ZVersion = WWVersion
                        ZCliente = WWCliente
                        ZNombre = WWDesCliente
                        ZFecha = WWFecha
                        ZFechaent = WWFecEntrega
                        ZTipoPedido = WTipoPedido
                        ZCondicion = WWDesPago
                        ZEntrega = WWDirentrega
                        ZObservaciones1 = Left$(WObservaciones, 50)
                        ZObservaciones2 = Right$(WObservaciones, 50)
                        ZOrden = WWOrdenCpa
                        ZArticulo = ""
                        ZDescripcion = ""
                        ZPrecio = "0"
                        ZCantidad = "0"
                        ZEnvase = ""
                        
                        spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                                "Cantidad , Envase, Lote, CantiLote )" + _
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
                                "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                                "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                                
                        Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                Next Ciclo
                    
            End If
                
        End If
            
    Next A
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
        SumaEspe = SumaEspe + 1
    
        Auxi = WPedido
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = WPedido
        ZRenglon = Str$(WRenglon)
        ZEmpresa = ""
        ZVersion = WWVersion
        ZCliente = WWCliente
        ZNombre = WWDesCliente
        ZFecha = WWFecha
        ZFechaent = WWFecEntrega
        ZTipoPedido = WTipoPedido
        ZCondicion = WWDesPago
        ZEntrega = WWDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = WWOrdenCpa
        ZArticulo = ""
        ZDescripcion = ""
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
        ZZLote = ""
        ZZCantiLote = ""
                        
        spImprePedIp = "INSERT INTO ImprePedIpPigmento (" + _
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
                    "Cantidad , Envase, Lote, CantiLote )" + _
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
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                    "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                                
        Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePedIpPigmento SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "',"
    ZSql = ZSql + "TipoPed = " + "'" + Str$(WWTipoped) + "'"
    spImprePedIp = ZSql
    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePedIpPigmento ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "ImprePedsqlipPigmento.rpt"
    
    Listado.Destination = 1
  Rem   Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
    
    
    ZZRequiereCertificado = ""
    ZZRequiereMsds = ""
    ZZRequiereMsdsCada = ""
    ZZRequiereHoja = ""
    ZZPermiteParcial = ""
    ZZPartidasVarias = ""

    ZZEmailCertificado = ""
    ZZEmailMsds = ""
    ZZEmailHoja = ""
    ZZDiasI = ""
    ZZDiasII = ""
    ZZDiasIII = ""
    ZZEnvasesI = ""
    ZZEnvasesII = ""
    ZZEnvasesIII = ""
    ZZEtiquetaI = ""
    ZZEtiquetaII = ""
    ZZEspecif1 = ""
    ZZEspecif2 = ""
    ZZEspecif3 = ""
    ZZEspecif4 = ""
    ZZEspecif5 = ""
    ZZCantidadPartidas = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteEspecif"
    ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + WWCliente + "'"
    spClienteEspecif = ZSql
    Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstClienteEspecif.RecordCount > 0 Then
    
        ZZRequiereCertificado = IIf(IsNull(rstClienteEspecif!RequiereCertificado), "0", rstClienteEspecif!RequiereCertificado)
        ZZRequiereMsds = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
        ZZRequiereMsdsCada = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
        ZZRequiereHoja = IIf(IsNull(rstClienteEspecif!RequiereHoja), "0", rstClienteEspecif!RequiereHoja)
        ZZPermiteParcial = IIf(IsNull(rstClienteEspecif!PermiteParcial), "0", rstClienteEspecif!PermiteParcial)
        ZZPartidasVarias = IIf(IsNull(rstClienteEspecif!PartidaVarias), "0", rstClienteEspecif!PartidaVarias)

        ZZEmailCertificado = IIf(IsNull(rstClienteEspecif!EmailCertificado), "", rstClienteEspecif!EmailCertificado)
        ZZEmailMsds = IIf(IsNull(rstClienteEspecif!EmailMsds), "", rstClienteEspecif!EmailMsds)
        ZZEmailHoja = IIf(IsNull(rstClienteEspecif!EmailHoja), "", rstClienteEspecif!EmailHoja)
        ZZDiasI = IIf(IsNull(rstClienteEspecif!DiasI), "", rstClienteEspecif!DiasI)
        ZZDiasII = IIf(IsNull(rstClienteEspecif!DiasII), "", rstClienteEspecif!DiasII)
        ZZDiasIII = IIf(IsNull(rstClienteEspecif!DiasIII), "", rstClienteEspecif!DiasIII)
        ZZEnvasesI = IIf(IsNull(rstClienteEspecif!EnvasesI), "", rstClienteEspecif!EnvasesI)
        ZZEnvasesII = IIf(IsNull(rstClienteEspecif!EnvasesII), "", rstClienteEspecif!EnvasesII)
        ZZEnvasesIII = IIf(IsNull(rstClienteEspecif!EnvasesIII), "", rstClienteEspecif!EnvasesIII)
        ZZEtiquetaI = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
        ZZEtiquetaII = IIf(IsNull(rstClienteEspecif!EtiquetaII), "", rstClienteEspecif!EtiquetaII)
        ZZEspecif1 = IIf(IsNull(rstClienteEspecif!Especif1), "", rstClienteEspecif!Especif1)
        ZZEspecif2 = IIf(IsNull(rstClienteEspecif!Especif2), "", rstClienteEspecif!Especif2)
        ZZEspecif3 = IIf(IsNull(rstClienteEspecif!Especif3), "", rstClienteEspecif!Especif3)
        ZZEspecif4 = IIf(IsNull(rstClienteEspecif!Especif4), "", rstClienteEspecif!Especif4)
        ZZEspecif5 = IIf(IsNull(rstClienteEspecif!Especif5), "", rstClienteEspecif!Especif5)
        ZZCantidadPartidas = IIf(IsNull(rstClienteEspecif!CantidadPartidas), "", rstClienteEspecif!CantidadPartidas)
        
        rstClienteEspecif.Close
        
    End If
    
    ZZImprime = "N"
    
    If Val(ZZRequiereCertificado) <> 0 Or Val(ZZRequiereMsds) <> 0 Or Val(ZZRequiereMsdsCada) <> 0 Or Val(ZZRequiereHoja) <> 0 Or Val(ZZPermiteParcial) <> 0 Or Val(ZZPartidasVarias) <> 0 Then
        ZZImprime = "S"
    End If
    If Trim(ZZDiasI) <> "" Or Trim(ZZDiasII) <> "" Or Trim(ZZDiasIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEnvasesI) <> "" Or Trim(ZZEnvasesII) <> "" Or Trim(ZZEnvasesIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEtiquetaI) <> "" Or Trim(ZZEtiquetaII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEspecif1) <> "" Or Trim(ZZEspecif2) <> "" Or Trim(ZZEspecif3) <> "" Or Trim(ZZEspecif4) <> "" Or Trim(ZZEspecif5) <> "" Then
        ZZImprime = "S"
    End If
    If Val(ZZCantidadPartidas) <> 0 Then
        ZZImprime = "S"
    End If
    
    If ZZImprime = "S" Then
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT ImprePed.Clave, ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.TipoPedido, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via, " _
                + "ClienteEspecif.RequiereCertificado, ClienteEspecif.RequiereMsds, ClienteEspecif.RequiereMsdsCada, ClienteEspecif.RequiereHoja, ClienteEspecif.PermiteParcial, ClienteEspecif.DiasI, ClienteEspecif.DiasII, ClienteEspecif.DiasIII, ClienteEspecif.Especif1, ClienteEspecif.Especif2, ClienteEspecif.Especif3, ClienteEspecif.Especif4, ClienteEspecif.Especif5, ClienteEspecif.PartidaVarias, ClienteEspecif.CantidadPartidas, ClienteEspecif.EnvasesI, ClienteEspecif.EnvasesII, ClienteEspecif.EnvasesIII, ClienteEspecif.EtiquetaI, ClienteEspecif.EtiquetaII " _
                + "From " _
                + DSQ + ".dbo.ImprePedIpPigmento ImprePed, " _
                + DSQ + ".dbo.ClienteEspecif ClienteEspecif " _
                + "Where " _
                + "ImprePed.Cliente = ClienteEspecif.Cliente AND " _
                + "ImprePed.Pedido >= 0 AND " _
                + "ImprePed.Pedido <= 999999"
                            
        Listado.Connect = Connect()
        Listado.ReportFileName = "ImprePedsqlEspecifIpPigmento.rpt"
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.CopiesToPrinter = 1
        Listado.Action = 1
        
    End If
        
    Exit Sub
        
WError:
    Resume Next

End Sub





