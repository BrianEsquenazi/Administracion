VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImprepedExpe 
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
Attribute VB_Name = "PrgImprepedExpe"
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
Dim Auxiliar(100, 4) As String
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
Dim ZZLugarDirEntrega As Integer
Dim ZZDirEntrega(10) As String
Dim WWEspecif(100) As String
Dim WWRazon As String
Dim WWPago As String
Dim WWDirentrega As String
Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As Double
Dim WWPrecio As Double
Dim WWObserva As String
Dim WWOrdenCpa As String
Dim WWDesPago As String
Dim WWVia As String
Dim WWCliente As String


Dim WWVersion As String
Dim WWDesCliente As String
Dim WWFecha As String
Dim WWFecEntrega As String
Dim WWObservaciones As String
Dim WWTipoped As Integer

Dim ZZRequiereCertificado As String
Dim ZZRequiereMsds As String
Dim ZZRequiereMsdsCada As String
Dim ZZRequiereHoja As String
Dim ZZPermiteParcial As String
Dim ZZPartidasVarias As String

Dim ZZEmailCertificado As String
Dim ZZEmailMsds As String
Dim ZZEmailHoja As String
Dim ZZDiasI As String
Dim ZZDiasII As String
Dim ZZDiasIII As String
Dim ZZEnvasesI As String
Dim ZZEnvasesII As String
Dim ZZEnvasesIII As String
Dim ZZEtiquetaI As String
Dim ZZEtiquetaII As String
Dim ZZEspecif1 As String
Dim ZZEspecif2 As String
Dim ZZEspecif3 As String
Dim ZZEspecif4 As String
Dim ZZEspecif5 As String
Dim ZZCantidadPartidas As String
Dim ImpreEnvase(10) As String

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

    spPedido = "ListaPedidoTotalListado4"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
                If rstPedido!TipoPedido = 5 Then
                
                Entra = "S"
                            
                For XDa = 1 To Lugar
                    If Vector(Lugar, 1) = rstPedido!PEDIDO Then
                        Entra = "N"
                        Exit For
                    End If
                Next XDa
                                
                If Entra = "S" Then
                    Lugar = Lugar + 1
                    Vector(Lugar, 1) = rstPedido!PEDIDO
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
    
    
    
    
    
    
    
    
    
    Rem Pasa = 0
    
    Rem Sql1 = "Select *"
    Rem Sql2 = " FROM Muestra"
    Rem Sql3 = " Where Autoriza = " + "'" + "S" + "'"
    Rem Sql4 = " and Impresion = " + "'" + "X" + "'"
    Rem Sql5 = " Order by Pedido, Codigo"
    Rem spMuestra = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Rem Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMuestra.RecordCount > 0 Then
    Rem     With rstMuestra
    Rem         .MoveFirst
    Rem         If .NoMatch = False Then
    Rem             Do
    Rem
    Rem                 XTipoPro = ""
    Rem
    Rem                 WTerminado = Trim(rstMuestra!Producto)
    Rem                 WArticulo = Trim(rstMuestra!Articulo)
    Rem
    Rem                 If WTerminado <> "" Then
    Rem                     XCodigo = Val(Mid$(WTerminado, 4, 5))
    Rem                     If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DW" Then
    Rem                         XTipoPro = "CO"
    Rem                             Else
    Rem                         If XCodigo >= 0 And XCodigo <= 999 Then
    Rem                             XTipoPro = "CO"
    Rem                                 Else
    Rem                             If XCodigo >= 11000 And XCodigo <= 11999 Then
    Rem                                 XTipoPro = "CO"
    Rem                                     Else
    Rem                                 If XCodigo >= 25000 And XCodigo <= 25999 Then
    Rem                                     XTipoPro = "FA"
    Rem                                         Else
    Rem                                     If XCodigo >= 2300 And XCodigo <= 2399 Then
    Rem                                         XTipoPro = "BI"
    Rem                                             Else
    Rem                                         XTipoPro = "PT"
    Rem                                     End If
    Rem                                 End If
    Rem                             End If
    Rem                         End If
    Rem                     End If
    Rem                 End If
    Rem
    Rem                 If WArticulo <> "" Then
    Rem                    If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Or Left$(WArticulo, 2) = "CO" Then
    Rem                         XTipoPro = "CO"
    Rem                             Else
    Rem                         XTipoPro = "PT"
    Rem                     End If
    Rem                 End If
    Rem
    Rem                 If XTipoPro = "CO" Then
    Rem
    Rem                 If Pasa = 0 Then
    Rem                     Pasa = 1
    Rem                     Lugar = Lugar + 1
    Rem                     Vector(Lugar, 1) = rstMuestra!PEDIDO
    Rem                     Vector(Lugar, 2) = "3"
    Rem                     Corte = rstMuestra!PEDIDO
    Rem                 End If
    Rem                 If Corte <> rstMuestra!PEDIDO Then
    Rem                     Lugar = Lugar + 1
    Rem                     Vector(Lugar, 1) = rstMuestra!PEDIDO
    Rem                     Vector(Lugar, 2) = "3"
    Rem                     Corte = rstMuestra!PEDIDO
    Rem                 End If
    Rem
    Rem                 End If
    Rem
    Rem                 .MoveNext
    Rem                 If .EOF = True Then
    Rem                     Exit Do
    Rem                 End If
    Rem             Loop
    Rem         End If
    Rem     End With
    Rem     rstMuestra.Close
    Rem End If
    
    
    
    
    
    Rem Lugar = 1
    Rem Vector(Lugar, 1) = "377914"
    Rem Vector(Lugar, 2) = "1"
    Rem Vector(Lugar, 3) = "0"
    
    
    
    
    
    If Lugar > 0 Then
        PrgImprepedExpe.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImprepedExpe.Refresh
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
    
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        WTipoped = Vector(WWCicla, 3)
        
        Select Case Val(WTipoPedido)
            Case 1
                Call Proceso_Click
                Call ImpresionSql
                Rem If Val(WTipoped) = 5 Then
                Rem     Call ImpresionIII
                Rem         Else
                Rem     Call Impresion
                Rem End If
                
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
                
            Case 2
                Rem Call ProcesoII_click
                Rem Call ImpresionII
                Rem WMarca = "X"
                                
                Rem XParam = "'" + WPedido + "','" _
                rem              + WMarca + "'"
                Rem
                Rem spPedido = "ModificaPedidoImpresion1 " + XParam
                Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                Rem Call ProcesoIII_click
                Rem Call ImpresionIII
                Rem WMarca = "S"
                        
                Rem Sql1 = "UPDATE Muestra SET "
                Rem Sql2 = " Impresion =  " + "'" + WMarca + "'"
                Rem Sql3 = " Where Pedido = " + "'" + WPedido + "'"
                Rem spMuestra = Sql1 + Sql2 + Sql3
                Rem Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
                           
        End Select
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImprepedExpe.Hide
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
                    
                    If WObserva <> "" Then
                        Print #1, Tab(1); "|";
                        Print #1, Tab(16); "|Observ.:";
                        Print #1, WObserva;
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
                        WVia = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                        
                        WWCliente = rstPedido!Cliente
                        WWFecha = rstPedido!Fecha
                        WWFecEntrega = rstPedido!FecEntrega
                        WWVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                        WWTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                        WWObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                        WWObservaciones = Left$(WObservaciones + Space$(100), 100)
                        ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                        WWVia = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                        WWOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Rem Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad)
                        Datos(Renglon, 8) = IIf(IsNull(rstPedido!observa), "", rstPedido!observa)
                        Datos(Renglon, 9) = IIf(IsNull(rstPedido!Especificaciones), "", rstPedido!Especificaciones)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!terminado
                
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
        terminado = Auxiliar(da, 2)
        
        spPrecios = "ConsultaPrecios " + "'" + Cliente + terminado + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Renglon = Renglon + 1
            Datos(Renglon, 1) = rstPrecios!Descripcion
            Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
            rstPrecios.Close
        End If
        
    Next da

End Sub


Private Sub Procesoanterior_Click()

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
                        
                        Datos(Renglon, 0) = rstPedido!terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
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
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!terminado
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
        terminado = Auxiliar(da, 2)
        Tipopro = Auxiliar(da, 3)
        Articulo = Auxiliar(da, 4)
        
        Renglon = Renglon + 1
        
        If Left$(terminado, 2) <> "DY" And Left$(terminado, 2) <> "DW" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + terminado + "'"
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
        WRazon = rstCliente!razon
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
                    If Val(xLote(da, Ciclo)) = 0 Then
                        xLote(da, Ciclo) = ""
                            Else
                        XParam = "'" + xLote(da, Ciclo) + "'"
                        spLaudo = "ListaLaudo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            xLote(da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
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
        
    For WDa = XLinea To 10
        Print #1, Tab(1); "|";
        Print #1, Tab(16); "|";
        Rem Print #1, Tab(47); "|";
        Print #1, Tab(47); "|";
        Print #1, Tab(53); "|";
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
                        
                        Datos(Renglon, 0) = rstPedido!terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        Datos(Renglon, 9) = IIf(IsNull(rstPedido!Especificaciones), "", rstPedido!Especificaciones)
                        
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
                        Auxiliar(WRenglon, 2) = rstPedido!terminado
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
        terminado = Auxiliar(da, 2)
        Tipopro = Auxiliar(da, 3)
        Articulo = Auxiliar(da, 4)
        
        Renglon = Renglon + 1
        
        If Left$(terminado, 2) <> "DY" And Left$(terminado, 2) <> "DW" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + terminado + "'"
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




Private Sub ImpresionSql()

    Rem On Error GoTo WError
    
    spImprePedIp = "Delete ImprePedIp"
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
    Select Case WWVia
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
    
    
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        WWPago = rstCliente!Pago1
        WWDesCliente = rstCliente!razon
        WWDirentrega = ""
        
        ZZDirEntrega(1) = rstCliente!DirEntrega
        ZZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WWDirentrega = ZZDirEntrega(ZZLugarDirEntrega)
        
        rstCliente.Close
        
        spPago = "ConsultaPago " + "'" + WWPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WWDesPago = rstPago!Nombre
            rstPago.Close
        End If
        
    End If
    
    
    
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For a = 1 To 40
        
        WCounter = WCounter + 1
                
        If Datos(WCounter, 0) <> "" Then
                
            WArticulo = Datos(WCounter, O)
            WDescripcion = Datos(WCounter, 1)
            WCantidad = Val(Datos(WCounter, 2))
            WPrecio = Val(Datos(WCounter, 3))
            WObserva = Datos(WCounter, 8)
            WEspecificaciones = Datos(WCounter, 9)
                
            Rem If WCantidad <> 0 Then
            
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
                
                ZArticulo = WArticulo
                ZDescripcion = WDescripcion
                ZPrecio = Str$(WPrecio)
                ZCantidad = Str$(WCantidad)
                ZEnvase = ImpreEnvase(1)
                ZZLote = ""
                ZZCantiLote = ""
                
                spImprePedIp = "INSERT INTO ImprePedIp (" + _
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
                    ZZLote = ""
                    ZZCantiLote = ""
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
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
                    ZZLote = ""
                    ZZCantiLote = ""
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
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
                    ZZLote = ""
                    ZZCantiLote = ""
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
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
                    
            Rem End If
                
        End If
            
    Next a
    
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
                        
        spImprePedIp = "INSERT INTO ImprePedIp (" + _
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
    ZSql = ZSql + "UPDATE ImprePedIp SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "',"
    ZSql = ZSql + "TipoPed = " + "'" + Str$(WWTipoped) + "'"
    spImprePedIp = ZSql
    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePedIp ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "ImprePedsqlip.rpt"
    Listado.Destination = 1
    Rem Listado.Destination = 0
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
        ZZEtiquetaII = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
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
                + DSQ + ".dbo.ImprePedIp ImprePed, " _
                + DSQ + ".dbo.ClienteEspecif ClienteEspecif " _
                + "Where " _
                + "ImprePed.Cliente = ClienteEspecif.Cliente AND " _
                + "ImprePed.Pedido >= 0 AND " _
                + "ImprePed.Pedido <= 999999"
                            
        Listado.Connect = Connect()
        Listado.ReportFileName = "ImprePedsqlEspecifIp.rpt"
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.CopiesToPrinter = 1
        Listado.Action = 1
        
    End If
        
    Exit Sub
        
WError:
    Resume Next

End Sub






