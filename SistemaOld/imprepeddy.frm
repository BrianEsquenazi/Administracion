VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImprepedDy 
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
         Caption         =   "Pedidos a Verificar"
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
Attribute VB_Name = "PrgImprepedDy"
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
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String
Dim Vector(1000, 3) As String
Dim Datos(100, 10) As String
Dim WPedido As String
Dim WCliente As String
Dim WVia As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WDespago As String
Dim WFecha As String
Dim WFecEntrega As String
Dim WVersion As String
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
Dim WEspecif(100) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

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

    Rem spPedido = "ListaPedidoTotalListado"
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    Rem
    Rem With rstPedido
    Rem
    Rem     .MoveFirst
    Rem     If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             Rem If rstPedido!Renglon = 1 Then
    Rem
    Rem                 dada = rstPedido!PEDIDO
    Rem
    Rem                 If rstPedido!Autorizo = "X" Then
    Rem
    Rem
    Rem                     If rstPedido!Impresion <> "X" Then
    Rem
    Rem                         XProducto = Val(Mid$(rstPedido!Terminado, 4, 5))
    Rem                         If XProducto < 25000 Or XProducto > 25999 Then
    Rem                         If XProducto < 2300 Or XProducto > 2399 Then
    Rem
    Rem                             XProducto1 = Left$(rstPedido!Terminado, 2)
    Rem
    Rem                             XCodigo = Val(Mid$(rstPedido!Terminado, 4, 5))
    Rem                             If Left$(rstPedido!Terminado, 2) = "DY" Then
    Rem                                 XTipoPro = "CO"
    Rem                                     Else
    Rem                                If XCodigo >= 0 And XCodigo <= 999 Then
    Rem                                     XTipoPro = "CO"
    Rem                                         Else
    Rem                                     If XCodigo >= 11000 And XCodigo <= 11999 Then
    Rem                                         XTipoPro = "CO"
    Rem                                             Else
    Rem                                         XTipoPro = "PT"
    Rem                                     End If
    Rem                                 End If
    Rem                             End If
    Rem
    Rem                             If XTipoPro = "CO" Then
    Rem
    Rem                                 Entra = "S"
    Rem
    Rem                                 For XDa = 1 To Lugar
    Rem                                     If Vector(Lugar, 1) = rstPedido!PEDIDO Then
    Rem                                         Entra = "N"
    Rem                                         Exit For
    Rem                                     End If
    Rem                                 Next XDa
    Rem
    Rem                                 If Entra = "S" Then
    Rem                                     Lugar = Lugar + 1
    Rem                                     Vector(Lugar, 1) = rstPedido!PEDIDO
    Rem                                     Vector(Lugar, 2) = "1"
    Rem                                 End If
    Rem
    Rem                             End If
    Rem
    Rem                         End If
    Rem
    Rem                     End If
    Rem                     End If
    Rem
    Rem                 End If
    Rem
    Rem             Rem End If
    Rem
    Rem             .MoveNext
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem         Loop
    Rem     End If
    Rem
    Rem End With
    Rem
    Rem End If
    
    
    
    
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
    
                    If XTipoPro = "CO" Then
    
                    If Pasa = 0 Then
                        Pasa = 1
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!PEDIDO
                        Vector(Lugar, 2) = "3"
                        Vector(Lugar, 3) = rstPedido!Tipoped
                        Corte = rstMuestra!PEDIDO
                    End If
                    If Corte <> rstMuestra!PEDIDO Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!PEDIDO
                        Vector(Lugar, 2) = "3"
                        Vector(Lugar, 3) = rstPedido!Tipoped
                        Corte = rstMuestra!PEDIDO
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
        PrgImprepedDy.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImprepedDy.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    If Val(WEmpresa) = 1 Then
        Rem Open "dada.txt" For Output As #1
        Open "lpt1" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
            Else
        Open "dada.txt" For Output As #1
        Rem Open "lpt1" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "11" + Chr$(72);
    End If

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        WTipoped = Vector(WWCicla, 3)
        
        Select Case Val(WTipoPedido)
            Case 1
                WPedido = Vector(WWCicla, 1)
                Call Proceso_Click
                If Val(WTipoped) = 5 Then
                    Call ImpresionIII
                        Else
                    Call Impresion
                End If
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
                
            Case Else
                Call ProcesoIII_click
                Call ImpresionIII
                WMarca = "S"
                        
                Sql1 = "UPDATE Muestra SET "
                Sql2 = " Impresion =  " + "'" + WMarca + "'"
                Sql3 = " Where Pedido = " + "'" + WPedido + "'"
                spMuestra = Sql1 + Sql2 + Sql3
                Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                    
        End Select
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImprepedDy.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
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
    
    ZVia = ""
    Select Case Val(WVia)
        Case 1
            ZVia = "Pedido Exportacion : " + "Terrestre"
        Case 2
            ZVia = "Pedido Exportacion : " + "Maritimo"
        Case 3
            ZVia = "Pedido Exportacion : " + "Aereo"
        Case Else
    End Select

    

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
        Print #1, WVersion; "  "; ZVia;
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
                WEspecifica = Datos(WCounter, 6)
                
                If Left$(WArticulo, 2) <> "DY" And Left$(WArticulo, 2) <> "DW" Then
                    WPartida = ""
                        Else
                    Call Busca_Partida
                End If
                    
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
                    
                    If Trim(WEspecifica) <> "" Then
                        Print #1, Tab(1); "|";
                        Print #1, Tab(16); "|Especif.:";
                        Print #1, WEspecifica;
                        Print #1, Tab(80); "|"
                        XLinea = XLinea + 1
                    End If
                    
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
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
                
       Print #1, Tab(1); "| Preparó: ";
          Print #1, Tab(16); "|     Etiquetó";
          Print #1, Tab(47); "|";
          Print #1, Tab(58); "|";
          Print #1, Tab(67); "|";
          Print #1, Tab(80); "|"
          Print #1, Tab(1); "| Fraccionó";
          Print #1, Tab(17); "|     Supervisó";
          Print #1, Tab(47); "| Despachó";
          Print #1, Tab(58); "|";
          Print #1, Tab(67); "|";
          Print #1, Tab(80); "|"
            
                
                
                
                
                
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
                        WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                        WTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                        WObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                        WObservaciones = Left$(WObservaciones + Space$(100), 100)
                        ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                        WVia = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!tipopro), "", rstPedido!tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        Datos(Renglon, 6) = IIf(IsNull(rstPedido!Especificaciones), "", rstPedido!Especificaciones)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!Terminado
                        Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!tipopro), "", rstPedido!tipopro)
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
        tipopro = Auxiliar(da, 3)
        Articulo = Auxiliar(da, 4)
        
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

Private Sub ImpresionIII()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
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
            WEspecifica = Datos(WCounter, 6)
            
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
                    
            If Trim(WEspecifica) <> "" Then
                Print #1, Tab(1); "|";
                Print #1, Tab(16); "|Especif.:";
                Print #1, WEspecifica;
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
    Sql3 = " Where Muestra.Pedido = " + "'" + Str$(WPedido) + "'"
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
                    WRazon = Trim(rstMuestra!Razon)
                    
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

