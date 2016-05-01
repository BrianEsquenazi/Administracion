VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgIpPelliTotal 
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
Attribute VB_Name = "PrgIpPelliTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim wtime As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String
Dim Vector(1000, 2) As String
Dim Datos(100, 13) As String
Dim WPedido As String
Dim WCliente As String
Dim WVia As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WDespago As String
Dim WFecha As String
Dim WFecEntrega As String
Dim WOrdenCpa As String
Dim WVersion As String
Dim WTipoped As String
Dim Lugar As Integer
Dim WEnvase(10) As String
Dim XEnvase(100, 6) As String
Dim XEspecificaciones(100) As String
Dim Auxiliar(1000, 2) As String
Dim WImpre(10) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim WEspecif(100) As String
Dim ImpreEnvase(10) As String

Private Sub Acepta_Click()

    WEmpresa = "0008"
    txtOdbc = "Empresa08"
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

    Sql1 = "Select *"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Autorizo = " + "'" + "X" + "'"
    Sql4 = " and Impresion3 = " + "'" + "M" + "'"
    Sql5 = " Order by Pedido"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
    
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                     If rstPedido!Renglon = 1 Then
                
                        If rstPedido!Autorizo = "X" Then
                    
                            If rstPedido!Impresion3 = "N" Then
                        
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
    
    
    
    
    
    
    
    Pasa = 0
    
    
    
    Rem Lugar = 1
    
    Rem Vector(Lugar, 1) = 5625
    Rem Vector(Lugar, 2) = "1"
    
    
    
    If Lugar > 0 Then
        PrgIpPelliTotal.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgIpPelliTotal.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

     For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        WTipoPedido = Vector(WWCicla, 2)
        Select Case Val(WTipoPedido)
            Case 1
                WPedido = Vector(WWCicla, 1)
                Call Proceso_Click
                Call Impresion
        
                WMarca = "X"
                
                Sql1 = "UPDATE Pedido SET "
                Sql2 = " Impresion3 =  " + "'" + WMarca + "'"
                Sql3 = " Where Pedido = " + "'" + WPedido + "'"
                spPedido = Sql1 + Sql2 + Sql3
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                           
        End Select
                
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgIpPelliTotal.Hide
    Unload Me
    End
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
                        WOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                Rem by nan
                         wtime = IIf(IsNull(rstPedido!ttime), "", rstPedido!ttime)
                         Rem by nan
                
                
                
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad)
                        Datos(Renglon, 11) = IIf(IsNull(rstPedido!Partida), "", rstPedido!Partida)
                        Datos(Renglon, 12) = IIf(IsNull(rstPedido!Observa), "", rstPedido!Observa)
                        Datos(Renglon, 13) = IIf(IsNull(rstPedido!Proceso2), "", rstPedido!Proceso2)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!Terminado
                
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
        
        spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Renglon = Renglon + 1
            Datos(Renglon, 1) = rstPrecios!Descripcion
            Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
        End If
        
    Next da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Impresion()

    On Error GoTo WError
    
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
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
    

    
    
    WObservaciones = Left$(WObservaciones + Space$(100), 100)
    Select Case WTipoped
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
    
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
    
    For WCounter = 1 To 99
        
        If Datos(WCounter, 0) <> "" Then
                    
            WArticulo = Datos(WCounter, 0)
            WDescripcion = Datos(WCounter, 1)
            WCantidad = Val(Datos(WCounter, 2))
            WPrecio = Val(Datos(WCounter, 3))
            WEspecificaciones = Datos(WCounter, 4)
            WPartida = Datos(WCounter, 11)
            WObserva = Datos(WCounter, 12)
            WProceso2 = Datos(WCounter, 13)
            If Val(WProceso2) = 2 Then
                WTarea = "Produccion"
                    Else
                WTarea = ""
            End If
            
            WLugar = WCounter
                
            If WArticulo <> "" Then
            
                If WCantidad <> 0 Then
                
                    Erase ImpreEnvase
                    LugarEnvase = 0
                
                    For Cicla = 1 To 6 Step 2
                        If Val(XEnvase(WLugar, Cicla)) <> 0 Then
                            Rem If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                            LugarEnvase = LugarEnvase + 1
                            spEnvase = "ConsultaEnvases " + "'" + XEnvase(WLugar, Cicla) + "'"
                            Rem spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                            
                            If rstEnvase.RecordCount > 0 Then
                                WAbre = rstEnvase!Abreviatura
                                rstEnvase.Close
                                    Else
                                WAbre = ""
                            End If
                            ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WLugar, Cicla + 1))) + " " + Left$(WAbre, 8)
                            Rem ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8)
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
                    ZEmpresa = WNombreEmpresa
                    ZVersion = WVersion
                    ZCliente = WCliente
                    ZNombre = WRazon
                    ZFecha = WFecha
                    ZFechaent = WFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WDespago
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WOrdenCpa
                    ZArticulo = WArticulo
                    ZDescripcion = WDescripcion
                    ZPrecio = Str$(WPrecio)
                    ZCantidad = Str$(WCantidad)
                    ZEnvase = ImpreEnvase(1)
                    
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
                                "Descripcion , Precio , ttime ," + _
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
                                "'" + wtime + "'," + _
                                "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
                    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE ImprePed SET "
                    ZSql = ZSql + " Partida = " + "'" + WPartida + "',"
                    ZSql = ZSql + " Tarea = " + "'" + WTarea + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                    spImprePed = ZSql
                    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                    
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
                        ZVersion = WVersion
                        ZCliente = WCliente
                        ZNombre = WRazon
                        ZFecha = WFecha
                        ZFechaent = WFecEntrega
                        ZTipoPedido = WTipoPedido
                        ZCondicion = WDespago
                        ZEntrega = WDirentrega
                        ZObservaciones1 = Left$(WObservaciones, 50)
                        ZObservaciones2 = Right$(WObservaciones, 50)
                        ZOrden = WOrdenCpa
                        ZArticulo = ""
                        ZDescripcion = Left$("Obs.: " + WObserva, 50)
                        ZPrecio = "0"
                        ZCantidad = "0"
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
                    
                    End If
                    
                    
                    If Trim(WEspecificaciones) <> "" And Trim(WEspecificaciones) <> "0" Then
                    
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
                        ZVersion = WVersion
                        ZCliente = WCliente
                        ZNombre = WRazon
                        ZFecha = WFecha
                        ZFechaent = WFecEntrega
                        ZTipoPedido = WTipoPedido
                        ZCondicion = WDespago
                        ZEntrega = WDirentrega
                        ZObservaciones1 = Left$(WObservaciones, 50)
                        ZObservaciones2 = Right$(WObservaciones, 50)
                        ZOrden = WOrdenCpa
                        ZArticulo = ""
                        ZDescripcion = Left$("Especificaciones: " + WEspecificaciones, 50)
                        ZPrecio = "0"
                        ZCantidad = "0"
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
                    
                    End If
                    
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
                        ZEmpresa = WNombreEmpresa
                        ZVersion = WVersion
                        ZCliente = WCliente
                        ZNombre = WRazon
                        ZFecha = WFecha
                        ZFechaent = WFecEntrega
                        ZTipoPedido = WTipoPedido
                        ZCondicion = WDespago
                        ZEntrega = WDirentrega
                        ZObservaciones1 = Left$(WObservaciones, 50)
                        ZObservaciones2 = Right$(WObservaciones, 50)
                        ZOrden = WOrdenCpa
                       
                        
                        ZArticulo = ""
                        ZDescripcion = ""
                        ZPrecio = "0"
                        ZCantidad = "0"
                        ZEnvase = ImpreEnvase(Ciclo)
                        
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
                        
                End If
                    
            End If
                                        
        End If
            
    Next WCounter
    
    SumaEspe = 0
    
    
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
        ZEmpresa = WNombreEmpresa
        ZVersion = WVersion
        ZCliente = WCliente
        ZNombre = WRazon
        ZFecha = WFecha
        ZFechaent = WFecEntrega
        ZTipoPedido = WTipoPedido
        ZCondicion = WDespago
        ZEntrega = WDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = WOrdenCpa
        ZArticulo = ""
        ZDescripcion = Left$(WEspecif(SumaEspe), 50)
        If SumaEspe = 1 And Trim(ZDescripcion) <> "" Then
            ZDescripcion = Left$("Especif.:" + ZDescripcion, 50)
        End If
        ZPrecio = "0"
        ZCantidad = "0"
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
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePed SET "
    ZSql = ZSql + "Via = " + "'" + UCase(ZVia) + "'"
    spImprePed = ZSql
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    Rem If Tipoped.ListIndex = 5 Or Tipoped.ListIndex = 6 Then
    Rem     Listado.ReportFileName = "ImprepedsqlMuestra.rpt"
    Rem         Else
    Rem      Listado.ReportFileName = "Imprepedsql.rpt"
    Rem End If
    
    Listado.ReportFileName = "ImprepedsqlII.rpt"
    
    Listado.Destination = 0
    Rem  Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub



