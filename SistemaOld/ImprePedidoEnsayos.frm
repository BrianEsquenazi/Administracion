VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImprePedidoEnsayos 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos de Ensayos"
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
      Left            =   480
      Top             =   2280
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
Attribute VB_Name = "PrgImprePedidoEnsayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstImprePedII As Recordset
Dim spImprePedII As String
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
Dim WRenglon As Integer

Dim WEnvase(10) As String
Dim XEnvase(40, 6) As String
Dim Auxiliar(100, 5) As String
Dim WImpre(10) As String
Dim WArticulo As String
Dim WCantidad As Double
Dim WPartida As String
Dim WEspecif(100) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
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
    
    Sql1 = "Select Pedido, Fecha, Cliente, FecEntrega, TipoPed, Autorizo, Impresion, Cantidad, Facturado, Precio, Impresion1, Terminado, Proceso1, TipoPedido, Impresion2"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.Autorizo = " + "'" + "X" + "'"
    Sql4 = " and Pedido.Impresion = " + "'" + "X" + "'"
    Rem Sql5 = " and Pedido.Impresion1 = " + "'" + "N" + "'"
    Sql6 = " Order by Pedido.Pedido"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Rem If rstPedido!pedido = 320348 Then Stop
            
                ZImpre = IIf(IsNull(rstPedido!impresion2), "", rstPedido!impresion2)
                
                If ZImpre <> "X" Then
                
                XProducto = Mid$(rstPedido!Terminado, 1, 2)
                If XProducto = "YQ" Or XProducto = "YF" Or XProducto = "YH" Then
                
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
    rstPedido.Close
    
    End If
    
    If Lugar > 0 Then
        PrgImprePedidoEnsayos.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImprePedidoEnsayos.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub


Private Sub Imprime_Click()

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        
        Call Proceso_Click
        Call Impresion
                
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + "Impresion2 = " + "'" + "X" + "'"
        ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Next WWCicla
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImprePedidoEnsayos.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    On Error GoTo WError
    
    spImprePedII = "Delete ImprePedII"
    Set rstImprePedII = db.OpenRecordset(spImprePedII, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Pedido = " + "'" + WPedido + "'"
    ZSql = ZSql + " Order by Pedido.Pedido"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZZCliente = rstPedido!Cliente
        ZZFecha = rstPedido!Fecha
        ZZFecEntrega = rstPedido!FecEntrega
        ZZVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
        ZZTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
        ZZObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
        ZZObservaciones = Left$(ZZObservaciones + Space$(100), 100)
        ZZOrdenCpa = rstPedido!ordencpa
        ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
        rstPedido.Close
    End If
    
    WObservaciones = Left$(ZZObservaciones + Space$(100), 100)
    Select Case ZZTipoped
        Case 0
            ZZTipoPedido = " (Normal)"
        Case 1
            ZZTipoPedido = " (A fecha)"
        Case 2
            ZZTipoPedido = " (Fecha Limite)"
        Case 3
            ZZTipoPedido = " (Urgente)"
        Case 4
            ZZTipoPedido = " (Retira Cliente)"
        Case 5
            ZZTipoPedido = " (Muestra)"
        Case Else
    End Select
    
    spCliente = "ConsultaCliente " + "'" + ZZCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZDesCliente = rstCliente!razon
        ZZPago = rstCliente!Pago1
        ZZDirentrega = ""
        ZDirEntrega(1) = rstCliente!DirEntrega
        ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        WDirentrega = ZDirEntrega(ZZLugarDirEntrega)
                
        Erase WEspecif
                
        WEspecif(1) = ""
        WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
        WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
        WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
        WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
        WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
        WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
        WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
        WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
        WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
        For CicloEspecif = 1 To 10
            WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
        Next CicloEspecif
                
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + Str$(ZZPago) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            ZZDesPago = rstPago!Nombre
            rstPago.Close
        End If
        
    End If
    
    ZZRenglon = 0
    
    For Ciclo = 1 To WRenglon
    
        ZZTerminado = Datos(Ciclo, 0)
    
        If ZZTerminado <> "" Then
                    
            WArticulo = Datos(Ciclo, 0)
            WDescripcion = Datos(Ciclo, 1)
            WCantidad = Val(Datos(Ciclo, 2))
            WPrecio = Val(Datos(Ciclo, 3))
            WEspecificaciones = WEspecif(Ciclo)
            
            Rem If WCantidad <> 0 Then
                
                Erase ImpreEnvase
                LugarEnvase = 0
                
                For Cicla = 1 To 6 Step 2
                    If Val(XEnvase(WLugar, Cicla)) <> 0 Then
                        LugarEnvase = LugarEnvase + 1
                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WLugar, Cicla) + "'"
                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnvase.RecordCount > 0 Then
                            WAbre = rstEnvase!Abreviatura
                            rstEnvase.Close
                                Else
                            WAbre = ""
                        End If
                        ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WLugar, Cicla + 1))) + " " + Left$(WAbre, 8)
                    End If
                Next Cicla
                
                ZZRenglon = ZZRenglon + 1
                    
                Auxi = WPedido
                Call Ceros(Auxi, 6)
                Auxi1 = ZZRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = WPedido
                ZRenglon = Str$(ZZRenglon)
                ZEmpresa = ""
                ZVersion = Str$(ZZVersion)
                ZCliente = ZZCliente
                ZNombre = ZZZDesCliente
                ZFecha = ZZFecha
                ZFechaent = ZZFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = ZZDesPago
                ZEntrega = WDirentrega
                ZObservaciones1 = Left$(ZZObservaciones, 50)
                ZObservaciones2 = Right$(ZZObservaciones, 50)
                ZOrden = ZZOrdenCpa
                ZArticulo = WArticulo
                ZDescripcion = WDescripcion
                ZPrecio = Str$(WPrecio)
                ZCantidad = Str$(WCantidad)
                ZEnvase = ImpreEnvase(1)
                    
                spImprePedII = "INSERT INTO ImprePedII (" + _
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
                                
                Set rstImprePedII = db.OpenRecordset(spImprePedII, dbOpenSnapshot, dbSQLPassThrough)
                
                If WEspecificaciones <> "" And WEspecificaciones <> "0" Then
                    
                    ZZRenglon = ZZRenglon + 1
                    
                    Auxi = WPedido
                    Call Ceros(Auxi, 6)
                    Auxi1 = ZZRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = WPedido
                    ZRenglon = Str$(WWRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = Str$(ZZVersion)
                    ZCliente = ZZCliente
                    ZNombre = ZZDesCliente
                    ZFecha = ZZFecha
                    ZFechaent = ZZFecEntrega
                    ZTipoPedido = ZZTipoPedido
                    ZCondicion = ZZDesPago
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = ZZOrdenCpa
                    ZArticulo = "Especif.:"
                    ZDescripcion = WEspecificaciones
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                        
                    spImprePedII = "INSERT INTO ImprePedII (" + _
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
                                
                    Set rstImprePedII = db.OpenRecordset(spImprePedII, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                    
                For CicloII = 2 To LugarEnvase
                    
                    ZZRenglon = ZZRenglon + 1
                    
                    Auxi = WPedido
                    Call Ceros(Auxi, 6)
                    Auxi1 = ZZRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = WPedido
                    ZRenglon = Str$(ZZRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = Str$(ZZVersion)
                    ZCliente = ZZCliente
                    ZNombre = ZZDesCliente
                    ZFecha = ZZFecha
                    ZFechaent = ZZFecEntrega
                    ZTipoPedido = ZZTipoPedido
                    ZCondicion = ZZDesPago
                    ZEntrega = WDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = ZZOrdenCpa
                    ZArticulo = ""
                    ZDescripcion = ""
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ImpreEnvase(CicloII)
                        
                    spImprePedII = "INSERT INTO ImprePedII (" + _
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
                                
                    Set rstImprePedII = db.OpenRecordset(spImprePedII, dbOpenSnapshot, dbSQLPassThrough)
                        
                Next CicloII
                        
            Rem End If
                    
        End If
            
    Next Ciclo
    
    SumaEspe = 0
    ZZDesde = ZZRenglon + 1
    
    For Ciclo = ZZDesde To 12
    
        ZZRenglon = ZZRenglon + 1
        SumaEspe = SumaEspe + 1
    
        Auxi = WPedido
        Call Ceros(Auxi, 6)
        Auxi1 = ZZRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = WPedido
        ZRenglon = Str$(ZZRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = Str$(ZZVersion)
        ZCliente = ZZCliente
        ZNombre = ZZDesCliente
        ZFecha = ZZFecha
        ZFechaent = ZZFecEntrega
        ZTipoPedido = ZZTipoPedido
        ZCondicion = ZZDesPago
        ZEntrega = WDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = ZZOrdenCpa
        ZArticulo = ""
        ZDescripcion = WEspecif(SumaEspe)
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
                        
        spImprePedII = "INSERT INTO ImprePedII (" + _
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
                                
        Set rstImprePedII = db.OpenRecordset(spImprePedII, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePedII.Pedido, ImprePedII.Version, ImprePedII.Cliente, ImprePedII.Nombre, ImprePedII.Fecha, ImprePedII.FechaEnt, ImprePedII.Condicion, ImprePedII.Entrega, ImprePedII.Observaciones1, ImprePedII.Observaciones2, ImprePedII.Orden, ImprePedII.ArticuloImprePedII.Descripcion, ImprePedII.Precio, ImprePedII.Cantidad, ImprePedII.Envase " _
            + "From " _
            + DSQ + ".dbo.ImprePedII ImprePedII " _
            + "Where " _
            + "ImprePedII.Pedido >= 0 AND ImprePedII.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "Imprepedsqlotro.rpt"
    Listado.Destination = 1
  Rem  Listado.Destination = 0
  Rem  Listado.CopiesToPrinter = 1
    Listado.Action = 1
        
    Exit Sub
        
WError:
    Resume Next


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
                    Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad)
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
        
        If ZZNombreComercial <> "" Then
            Datos(Renglon, 1) = ZZNombreComercial
            Datos(Renglon, 3) = Pusing("###,###.##", "0")
                Else
            spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Datos(Renglon, 1) = rstPrecios!Descripcion
                Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
                rstPrecios.Close
            End If
        End If
        
    Next Da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

