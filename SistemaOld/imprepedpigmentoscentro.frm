VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreped 
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
Attribute VB_Name = "PrgImpreped"
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
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstVendedor As Recordset
Dim spVendedor As String
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
Dim Auxiliar(100, 5) As String
Dim WImpre(10) As String
Dim WArticulo As String
Dim WCantidad As Double
Dim WPartida As String
Dim Partida(100, 3) As String
Dim LugarPartida As String
Dim WSaldo As Double
Dim xLote(10, 10) As String
Dim WCantiProceso As String
Dim WObserva As String
Dim WEspecif(100) As String
Dim WRazon As String
Dim WVendedor As String
Dim WDesVendedor As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

Dim Impre(100, 2) As String
Dim ImpreI(100, 2) As String
Dim ImpreII(100, 2) As String

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

    Rem spPedido = "ListaPedidoTotalListado4"
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Pedido.Proceso1, Pedido.TipoPedido, Pedido.Pedido, Pedido.TipoPed, Pedido.Impresion2, Pedido.autorizo "
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Impresion2 <> 'S'"
    ZSql = ZSql + " and Pedido.Autorizo = 'X'"
    ZSql = ZSql + " and Pedido.TipoPedido = 5"
    ZSql = ZSql + " Order by Pedido.Pedido"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                    If rstPedido!TipoPedido = 5 And rstPedido!Impresion2 <> "S" And rstPedido!autorizo = "X" Then
                    
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
    
    
    
    
    
    
    
    Rem spPedido = "ListaPedidoTotalListado2"
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
    Rem                 If rstPedido!Autorizo = "X" Then
    Rem
    Rem                     If rstPedido!Impresion1 = "N" Then
    Rem
    Rem                         Entra = "S"
    Rem
    Rem                         For XDa = 1 To Lugar
    Rem                             If Vector(Lugar, 1) = rstPedido!PEDIDO Then
    Rem                                 Entra = "N"
    Rem                                 Exit For
    Rem                             End If
    Rem                         Next XDa
    Rem
    Rem                         If Entra = "S" Then
    Rem                             Lugar = Lugar + 1
    Rem                             Vector(Lugar, 1) = rstPedido!PEDIDO
    Rem                             Vector(Lugar, 2) = "2"
    Rem                         End If
    Rem
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
    Rem rstPedido.Close
    Rem
    Rem End If
    
    
    
    
    
    
    
    
    
    Pasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Muestra.Autoriza, Muestra.Impresion, Muestra.Producto, Muestra.Articulo, Muestra.Ensayo, Muestra.Pedido"
    ZSql = ZSql + " FROM Muestra"
    ZSql = ZSql + " Where Autoriza = " + "'" + "S" + "'"
    ZSql = ZSql + " and Impresion = " + "'" + "X" + "'"
    ZSql = ZSql + " Order by Pedido, Codigo"
    spMuestra = ZSql
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        With rstMuestra
            .MoveFirst
            If .NoMatch = False Then
                Do
        
                    XTipoPro = ""
    
                    WTerminado = UCase(Trim(rstMuestra!Producto))
                    WArticulo = UCase(Trim(rstMuestra!Articulo))
                    WEnsayo = UCase(Trim(rstMuestra!Ensayo))
    
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
                    
                    If WEnsayo <> "" Then
                       If Left$(WEnsayo, 2) = "YF" Or Left$(WArticulo, 2) = "IF" Then
                            XTipoPro = "FA"
                                Else
                            XTipoPro = "PT"
                        End If
                    End If
    
                    If XTipoPro = "PT" Then
    
                    If Pasa = 0 Then
                        Pasa = 1
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!Pedido
                        Vector(Lugar, 2) = "3"
                        Corte = rstMuestra!Pedido
                    End If
                    If Corte <> rstMuestra!Pedido Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstMuestra!Pedido
                        Vector(Lugar, 2) = "3"
                        Corte = rstMuestra!Pedido
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
    
    
    
    
    
    Rem Lugar = 1
    Rem Vector(1, 1) = "350635"
    Rem Vector(1, 2) = "1"
    
    
    
    
    
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

Private Sub ImprimeAnterior_Click()

    Rem Open "dada.txt" For Output As #1
    Open "lpt1" For Output As #1

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        WTipoped = Vector(WWCicla, 3)
        
        Select Case Val(WTipoPedido)
            Case 1
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
                
                WMarca = "3"
                XParam = "'" + WPedido + "','" _
                             + WMarca + "'"
                                           
                spPedido = "ModificaPedidoProceso1 " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                A = 1
                
            Case 2
                Call ProcesoII_click
                Call ImpresionII
                WMarca = "X"
                                
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
                
                A = 2
                           
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
            rstPago.Close
        End If
    End If
    

    For XX = 1 To 1

        Print #1, Tab(1); String$(79, "-")
        
        Print #1, Tab(1); "|                         SURFACTAN S.A.";
        Print #1, Tab(80); "|"
        
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Tab(80); "|"
                
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
            Case 4
                Print #1, " (Retira Cliente)";
            Case 5
                Print #1, " (Muestra)";
            Case Else
        End Select
            
        Print #1, Tab(80); "|"
                
        Rem Print #1, Tab(1); "|";
        Rem Print #1, Tab(5); "C.Pago";
        Rem Print #1, Tab(28); ":";
        Rem Print #1, Tab(30); WDespago;
        Rem Print #1, Tab(80); "|"
                
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
                WObserva = Datos(WCounter, 8)
                
                Select Case WProceso2
                    Case 1
                        WPartida = "Stock"
                    Case 2
                        WPartida = "Pda:" + WCantiProceso
                    Case 3
                        WPartida = "Transfer."
                    Case 4
                        WPartida = "Prod. Pta (Fecha)"
                    Case 5
                        WPartida = "Parcial (Prod. Fecha)"
                    Case 6
                        WPartida = "Prod. P1 (MP)"
                    Case 7
                        WPartida = "FMP"
                    Case 8
                        WPartida = "Kgrs. (Pellital)"
                    Case 9
                        WPartida = "Varios"
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
                    Print #1, Tab(59); Pusing("###,###", Str$(WCantidad));
                    Print #1, Tab(67); "|";
                    
                    
                    
                    
                    
                    
                    
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    
                    
                    Erase Impre
                    Erase ImpreI
                    Erase ImpreII
                    XLugarI = 0
                    XLugarII = 0
                    XLugar = 0
                    
                    
                    XCanti = WCantidad
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja.Producto = " + "'" + WArticulo + "'"
                    ZSql = ZSql + " and Hoja.Saldo <> 0"
                    ZSql = ZSql + " and Hoja.Renglon = 1"
                    ZSql = ZSql + " Order by Hoja.Hoja"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                            
                        With rstHoja
                                                        
                            .MoveFirst
                            
                            If .NoMatch = False Then
                                Do
                            
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                            
                                    XLugarI = XLugarI + 1
                                    ImpreI(XLugarI, 1) = rstHoja!Hoja
                                    ImpreI(XLugarI, 2) = rstHoja!Saldo
                            
                                    .MoveNext
                            
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                
                                Loop
                            End If
                        End With
                        rstHoja.Close
                    End If
                        
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Terminado = " + "'" + WArticulo + "'"
                    ZSql = ZSql + " and Guia.Saldo <> 0"
                    ZSql = ZSql + " Order by Guia.Lote"
                    spGuia = ZSql
                    Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstGuia.RecordCount > 0 Then
                                
                        With rstGuia
                                
                            .MoveFirst
                                
                            If .NoMatch = False Then
                                Do
                                                            
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                    
                                    XLugarII = XLugarII + 1
                                    ImpreII(XLugarII, 1) = rstGuia!Lote
                                    ImpreII(XLugarII, 2) = rstGuia!Saldo
                                                
                                    .MoveNext
                                                            
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                
                                Loop
                            End If
                        End With
                        rstGuia.Close
                    End If
                    
                    ZLugarI = 0
                    ZLugarII = 0
                    ZLugar = 0
                    
                    ZZTotal = XLugarI + XLugarII
            
                    For ZCicla = 1 To ZZTotal
            
                        If ImpreI(ZLugarI + 1, 1) <> "" And ImpreII(ZLugarII + 1, 1) <> "" Then
                
                            If ImpreI(ZLugarI + 1, 1) < ImpreII(ZLugarII + 1, 1) Then
                
                                ZLugarI = ZLugarI + 1
                                ZLugar = ZLugar + 1
                                Impre(ZLugar, 1) = ImpreI(ZLugarI, 1)
                                Impre(ZLugar, 2) = ImpreI(ZLugarI, 2)
                        
                                    Else
                
                                ZLugarII = ZLugarII + 1
                                ZLugar = ZLugar + 1
                                Impre(ZLugar, 1) = ImpreII(ZLugarII, 1)
                                Impre(ZLugar, 2) = ImpreII(ZLugarII, 2)
                            
                            End If
                    
                                Else
                
                            If ImpreI(ZLugarI + 1, 1) <> "" Then
                                ZLugarI = ZLugarI + 1
                                ZLugar = ZLugar + 1
                                Impre(ZLugar, 1) = ImpreI(ZLugarI, 1)
                                Impre(ZLugar, 2) = ImpreI(ZLugarI, 2)
                            End If
                
                            If ImpreII(ZLugarII + 1, 1) <> "" Then
                                ZLugarII = ZLugarII + 1
                                ZLugar = ZLugar + 1
                                Impre(ZLugar, 1) = ImpreII(ZLugarII, 1)
                                Impre(ZLugar, 2) = ImpreII(ZLugarII, 2)
                            End If
                        
                        End If
                    
                    Next ZCicla
                    
                    
                    
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        
                    
                    
                                
                    For Cicla = 1 To 6 Step 2
                        If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                            Select Case Cicla
                                Case 1
                                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvase.RecordCount > 0 Then
                                        WAbre = rstEnvase!Abreviatura
                                        rstEnvase.Close
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
                                        rstEnvase.Close
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
                    
                    If Trim(WObserva) <> "" Then
                        Print #1, Tab(1); "|";
                        Print #1, Tab(16); "|Observ.:";
                        Print #1, WObserva;
                        Print #1, Tab(80); "|"
                        XLinea = XLinea + 1
                    End If
                    
                    Rem dada
                    If ZLugar < 4 Then
                    
                        Print #1, Tab(1); "|";
                        If Val(Impre(1, 2)) <> 0 Then
                            Print #1, Tab(3); "L:"; Pusing("######", Impre(1, 1));
                            Print #1, Tab(12); "Stk:"; Pusing("#####.##", Impre(1, 2));
                        End If
                        If Val(Impre(2, 2)) <> 0 Then
                            Print #1, Tab(30); "L:"; Pusing("######", Impre(2, 1));
                            Print #1, Tab(40); "Stk:"; Pusing("#####.##", Impre(2, 2));
                        End If
                        If Val(Impre(3, 2)) <> 0 Then
                            Print #1, Tab(58); "L:"; Pusing("######", Impre(3, 1));
                            Print #1, Tab(68); "Stk:"; Pusing("#####.##", Impre(3, 2));
                        End If
                        Print #1, Tab(80); "|"
                        XLinea = XLinea + 1
                        
                            Else
                            
                        Print #1, Tab(1); "|";
                        If Val(Impre(1, 2)) <> 0 Then
                            Print #1, Tab(3); "L:"; Pusing("######", Impre(1, 1));
                            Print #1, Tab(12); "Stk:"; Pusing("#####.##", Impre(1, 2));
                        End If
                        If Val(Impre(2, 2)) <> 0 Then
                            Print #1, Tab(30); "L:"; Pusing("######", Impre(2, 1));
                            Print #1, Tab(40); "Stk:"; Pusing("#####.##", Impre(2, 2));
                        End If
                        If Val(Impre(3, 2)) <> 0 Then
                            Print #1, Tab(58); "L:"; Pusing("######", Impre(3, 1));
                            Print #1, Tab(68); "Stk:"; Pusing("#####.##", Impre(3, 2));
                        End If
                        Print #1, Tab(80); "|"
                        XLinea = XLinea + 1
                        
                        Print #1, Tab(1); "|";
                        If Val(Impre(4, 2)) <> 0 Then
                            Print #1, Tab(3); "L:"; Pusing("######", Impre(4, 1));
                            Print #1, Tab(12); "Stk:"; Pusing("#####.##", Impre(4, 2));
                        End If
                        If Val(Impre(5, 2)) <> 0 Then
                            Print #1, Tab(30); "L:"; Pusing("######", Impre(5, 1));
                            Print #1, Tab(40); "Stk:"; Pusing("#####.##", Impre(5, 2));
                        End If
                        If Val(Impre(6, 2)) <> 0 Then
                            Print #1, Tab(58); "L:"; Pusing("######", Impre(6, 1));
                            Print #1, Tab(68); "Stk:"; Pusing("#####.##", Impre(6, 2));
                        End If
                        Print #1, Tab(80); "|"
                        XLinea = XLinea + 1
                    
                    End If
                    
                    Print #1, Tab(1); String$(79, "-")
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
        
        For WDa = XLinea To 14
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(47); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
        
        Rem agregado by nan
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "Preparó:      ";
        Print #1, "Etiquetó:       ";
        Print #1, "Fraccionó:      ";
        Print #1, "Supervisó:      ";
        Print #1, "Despachó:       "
                
        Rem Print #1, Tab(1); "|"; WImpre(1);
        Rem Print #1, Tab(10); "|"; WImpre(2);
        Rem Print #1, Tab(18); "|"; WImpre(3);
        Rem Print #1, Tab(26); "|"; WImpre(4);
        Rem Print #1, Tab(34); "|"; WImpre(5);
        Rem Print #1, Tab(42); "|"; WImpre(6);
        Rem Print #1, Tab(50); "|"; WImpre(7);
        Rem Print #1, Tab(58); "|"; WImpre(8);
        Rem Print #1, Tab(66); "|"; WImpre(9);
        Rem Print #1, Tab(80); "|"
        
        Rem Print #1, Tab(1); "|(020)";
        Rem Print #1, Tab(10); "|(021)";
        Rem Print #1, Tab(18); "|(022)";
        Rem Print #1, Tab(26); "|(023)";
        Rem Print #1, Tab(34); "|(024)";
        Rem Print #1, Tab(42); "|(025)";
        Rem Print #1, Tab(50); "|(026)";
        Rem Print #1, Tab(58); "|(030)";
        Rem Print #1, Tab(66); "|(028)";
        Rem Print #1, Tab(80); "|"
        Rem Print #1, Tab(1); String$(79, "-")
        
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
                        Datos(Renglon, 6) = IIf(IsNull(rstPedido!proceso2), "", rstPedido!proceso2)
                        Datos(Renglon, 7) = IIf(IsNull(rstPedido!cantiproceso), "", rstPedido!cantiproceso)
                        Datos(Renglon, 8) = IIf(IsNull(rstPedido!observa), "", rstPedido!observa)
                
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
                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                            Auxiliar(WRenglon, 5) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
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
    
        Cliente = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Tipopro = Auxiliar(Da, 3)
        Articulo = Auxiliar(Da, 4)
        ZZNombreComercial = Auxiliar(Da, 5)
        
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

End Sub

Private Sub ImpresionII()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
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
    Print #1, Tab(48); "Cant.";
    Print #1, Tab(53); "|";
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
                Print #1, Tab(47); "|";
                Print #1, Tab(48); Alinea("#####", Str$(WCantidad));
                Print #1, Tab(53); "|";
                
                For Ciclo = 1 To 9 Step 2
                    If Val(xLote(Da, Ciclo)) = 0 Then
                        xLote(Da, Ciclo) = ""
                            Else
                        XParam = "'" + xLote(Da, Ciclo) + "'"
                        spLaudo = "ListaLaudo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            xLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            rstLaudo.Close
                        End If
                    End If
                Next Ciclo
                    
                XLote1 = xLote(WLugar, 1)
                XLote2 = xLote(WLugar, 3)
                XLote3 = xLote(WLugar, 5)
                XLote4 = xLote(WLugar, 7)
                XLote5 = xLote(WLugar, 9)
                XCantiLote1 = xLote(WLugar, 2)
                XCantiLote2 = xLote(WLugar, 4)
                XCantiLote3 = xLote(WLugar, 6)
                XCantiLote4 = xLote(WLugar, 8)
                XCantiLote5 = xLote(WLugar, 10)
                
                
                                
                For Cicla = 1 To 10 Step 2
                    If Val(xLote(WCounter, Cicla)) <> 0 Then
                    
                        ZEnvase = ""
                        ZDescriEnvase = ""
                    
                        If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
                    
                            XArticulo = Left$(WArticulo, 3) + Right$(WArticulo, 7)
                        
                            spLaudo = "ListaLaudo " + "'" + xLote(WCounter, Cicla) + "'"
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WOrden = Str$(rstLaudo!Orden)
                                WPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                                rstLaudo.Close
                            End If
                    
                            XParam = "'" + WOrden + "','" _
                                    + XArticulo + "'"
                            spInforme = "ListaInformeOrdenArticulo " + XParam
                            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                            If rstInforme.RecordCount > 0 Then
                                ZEnvase = Str$(rstInforme!Envase)
                                rstInforme.Close
                            End If
            
                            spEnvase = "ConsultaEnvases " + "'" + ZEnvase + "'"
                            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                            If rstEnvase.RecordCount > 0 Then
                                ZDescriEnvase = rstEnvase!Abreviatura
                                rstEnvase.Close
                            End If
                        End If
                    
                        Select Case Cicla
                            Case 1
                                Print #1, Tab(54); Left$(WPartiOri, 10);
                                Print #1, Tab(64); Alinea("#####", xLote(WCounter, Cicla + 1));
                                Print #1, Tab(70); ZDescriEnvase;
                                Print #1, Tab(80); "|"
    
                            Case Else
                                Print #1, Tab(1); "|";
                                Print #1, Tab(16); "|";
                                Print #1, Tab(47); "|";
                                Print #1, Tab(53); "|";
                                Print #1, Tab(54); Left$(WPartiOri, 10);
                                Print #1, Tab(64); Alinea("#####", xLote(WCounter, Cicla + 1));
                                Print #1, Tab(70); ZDescriEnvase;
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
        
    For WDa = XLinea To 8
        Print #1, Tab(1); "|";
        Print #1, Tab(16); "|";
        Rem Print #1, Tab(47); "|";
        Print #1, Tab(47); "|";
        Print #1, Tab(53); "|";
        Print #1, Tab(80); "|"
    Next WDa
    Rem agregado by nan
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
                        
                        xLote(Renglon, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        xLote(Renglon, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        xLote(Renglon, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        xLote(Renglon, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        xLote(Renglon, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        xLote(Renglon, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        xLote(Renglon, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        xLote(Renglon, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        xLote(Renglon, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        xLote(Renglon, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                
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
    Next Da

End Sub

Private Sub ImpresionIII()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
        WVendedor = Str$(rstCliente!vendedor)
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
    Sql3 = " Where Muestra.Pedido = " + "'" + Str$(WPedido) + "'"
    Sql4 = " Order by Muestra.Codigo"
    spMuestra = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    If rstMuestra.RecordCount > 0 Then
        With rstMuestra
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WVendedor = Str$(rstMuestra!vendedor)
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
                        WLote = rstLaudo!Laudo
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


Private Sub Imprime_Click()
    PrgImpreped.Hide
    Unload Me
    PrgModifTerminado.Show
End Sub
