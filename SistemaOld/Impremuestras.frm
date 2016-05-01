VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreMuestras 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos de Muestras"
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
Attribute VB_Name = "PrgImpreMuestras"
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
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String
Dim WEnvase(10) As String
Dim XEnvase(100, 6) As String
Dim Auxiliar(1000, 2) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

Dim Vector(1000, 3) As String
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
Dim WVersion As String
Dim WTipoped As String
Dim Lugar As Integer
Dim WImpre(10) As String
Dim WEspecif(100) As String

Dim WWLote As String
Dim WWTipo As String

Dim AuxiliarII(100, 5) As String

Dim WWCliente As String
Dim WWFecha  As String
Dim WWFecEntrega As String
Dim WWVersion As String
Dim WWTipoped As String
Dim WWObservaciones As String

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
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where ImpreMuestra = " + "'" + "N" + "'"
    ZSql = ZSql + " and Terminado >= 'YQ-000-000'"
    ZSql = ZSql + " and Terminado <= 'YQ-999-999'"
    ZSql = ZSql + " Order by Clave"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                            
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
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstPedido.Close
        
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where ImpreMuestra = " + "'" + "N" + "'"
    ZSql = ZSql + " and Terminado >= 'YF-000-000'"
    ZSql = ZSql + " and Terminado <= 'YF-999-999'"
    ZSql = ZSql + " Order by Clave"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                            
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
        PrgImpreMuestras.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImpreMuestras.Refresh
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
        
        WPedido = Vector(WWCicla, 1)
        Call Proceso_Click
        Call ImpresionSql
        
        WMarca = "S"
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + " ImpreMuestra = " + "'" + WMarca + "'"
        ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreMuestras.Hide
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
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
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
    
    For Da = 1 To WRenglon
    
        Cliente = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        
        If Left$(Terminado, 2) = "DY" Or Left$(Terminado, 2) = "DW" Or Left$(Terminado, 2) = "DS" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                WClaveMp = Cliente + WArti
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Datos(Da, 3) = Pusing("###,###.##", rstPreciosMp!Precio)
                    rstPreciosMp.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Datos(Da, 1) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
        
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Datos(Da, 1) = rstPrecios!Descripcion
                    Datos(Da, 3) = Pusing("###,###.##", rstPrecios!Precio)
                    rstPrecios.Close
                End If
                
        End Select
        
    Next Da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub ImpresionSql()

    Rem On Error GoTo WError
    
    spImprePedip = "Delete ImprePedIp"
    Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
    
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
                    
    For A = 1 To 40
        
        WCounter = WCounter + 1
                
        If Datos(WCounter, 0) <> "" Then
                
            WArticulo = Datos(WCounter, O)
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
                
                spImprePedip = "INSERT INTO ImprePedIp (" + _
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
                            
                Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE ImprePedIp SET "
                ZSql = ZSql + " Lote = " + "'" + WPartida + "',"
                ZSql = ZSql + " Especificaciones = " + "'" + WTarea + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                spImprePedip = ZSql
                Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
                
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
                    
                    spImprePedip = "INSERT INTO ImprePedIp (" + _
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
                            
                    Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
                    
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
                    
                    spImprePedip = "INSERT INTO ImprePedIp (" + _
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
                            
                    Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
                
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
                    
                    spImprePedip = "INSERT INTO ImprePedIp (" + _
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
                            
                    Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                    
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
                        
        spImprePedip = "INSERT INTO ImprePedIp (" + _
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
                                
        Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePedIp SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "',"
    ZSql = ZSql + "TipoPed = " + "'" + WWTipoped + "'"
    spImprePedip = ZSql
    Set rstImprePedip = db.OpenRecordset(spImprePedip, dbOpenSnapshot, dbSQLPassThrough)
    
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




