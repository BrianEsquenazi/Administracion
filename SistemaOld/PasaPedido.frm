VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPasaPedido 
   Caption         =   "Traspaso de Pedidos de Pellital a Surfactan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ClientePelli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox ClienteSurfa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Cierre 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Ejecuta Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label DesClientePelli 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente Pellital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente Surfactan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label DesClienteSurfa 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "PrgPasaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPasaVector(100, 10) As String
Dim XEnvase(100, 6) As String
Dim XEspecificaciones(100) As String
Dim XDatosMuestra(100, 3) As String
Dim Auxiliar(100, 3) As String
Dim CargaEmpresa(12, 2) As String

Dim ZCliente As String
Dim ZTerminado As String
Dim ZArticulo As String
Dim ZClave As String
        
Dim ZPrecio As String
Dim ZDescripcion As String
        
Dim ZFecha1 As String
Dim ZFactura1 As String
Dim ZPrecio1 As String
Dim ZCantidad1 As String
        
Dim ZFecha2 As String
Dim ZFactura2 As String
Dim ZPrecio2 As String
Dim ZCantidad2 As String
        
Dim ZFecha3 As String
Dim ZFactura3 As String
Dim ZPrecio3 As String
Dim ZCantidad3 As String
        
Dim ZFecha4 As String
Dim ZFactura4 As String
Dim ZPrecio4 As String
Dim ZCantidad4 As String
        
Dim ZFecha5 As String
Dim ZFactura5 As String
Dim ZPrecio5 As String
Dim ZCantidad5 As String
        
Dim ZDate As String
Dim ZFecha As String
Dim ZPago As String
Dim ZEstado As String
Dim ZLugar As Integer

Dim ZZFecha As String
Dim ZZCliente As String
Dim ZZFecEntrega As String
Dim ZZHora As String
Dim ZZObservaciones As String
Dim ZZOrdenCpa As String
Dim ZZMarca1 As String
Dim ZZMarca2 As String
Dim ZZMarca3 As String
Dim ZZDestino As String
Dim ZZTipoped As String
Dim ZZVersion As String
Dim ZZLugarDirEntrega As String
Dim ZZVia As String


Dim ZZZZDescripcion As String
Dim ZZZZDescriEtiqueta As String
Dim ZZZZLinea As String
Dim ZZZZUnidad As String
Dim ZZZZInicial As String
Dim ZZZZEntradas As String
Dim ZZZZSalidas As String
Dim ZZZZMInimo As String
Dim ZZZZMinimo1 As String
Dim ZZZZFabrica As String
Dim ZZZZFabricaII As String
Dim ZZZZFabricaIII As String
Dim ZZZZDeposito As String
Dim ZZZZEnvase1 As String
Dim ZZZZEnvase2 As String
Dim ZZZZEnvase3 As String
Dim ZZZZEnvase4 As String
Dim ZZZZEnvase5 As String
Dim ZZZZEnvase6 As String
Dim ZZZZProceso As String
Dim ZZZZImpreadi As String
Dim ZZZZClase As String
Dim ZZZZSecundario As String
Dim ZZZZRiesgo As String
Dim ZZZZIntervencion As String
Dim ZZZZNaciones As String
Dim ZZZZEmbalaje As String
Dim ZZZZControla As String
Dim ZZZZSedronar As String
Dim ZZZZImpreVto As String
Dim ZZZZMarca As String
Dim ZZZZObservaciones As String
Dim ZZZZTipoeti As String
Dim ZZZZEscrito As String
Dim ZZZZPedido As String
Dim ZZZZConservacion As String
Dim ZZZZConservacionII As String
Dim ZZZZVida As String
Dim ZZZZSeguridad As String

Dim ZZZZVersion As String
Dim ZZZZVersionI As String
Dim ZZZZVersionII As String

Dim ZZZZFechaVersion As String
Dim ZZZZFechaVersionI As String
Dim ZZZZFechaVersionII As String

Dim ZZZZEstado As String
Dim ZZZZEstadoI As String
Dim ZZZZEstadoII As String

Dim ZZZZObserva As String
Dim ZZZZObservaI As String
Dim ZZZZObservaII As String

Dim ZZZZMetodo As String
Dim ZZZZEfluentes As String

Dim ZZZZCaracteristicas As String
Dim ZZZZCarga As String
Dim ZZZZEstadoProducto As String
Dim ZZZZListaProducto As String

Dim ZZZZDescripcionIngles As String
Dim ZZZZDescriEtiquetaIngles As String
Dim ZZZZConservacionIngles As String
Dim ZZZZConservacionIIIngles As String

Dim ZZZZResponsable As String

Dim ZZZZLoteAutorizado As String
Dim ZZZZCosto As String
Dim ZZZZFactor As String
Dim ZZZZDate As String

Dim ZZZZPrecioClave As String
Dim ZZZZPrecioPrecio As String
Dim ZZZZPrecioDescripcion As String
Dim ZZZZPrecioFecha As String
Dim ZZZZPrecioPago As String
Dim ZZZZPrecioEstado As String









Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Cierre_Click()
    PrgPasaPedido.Hide
    Unload Me
    PrgPedido.Show
End Sub

Private Sub Form_Load()
    Pedido.Text = ""
    ClienteSurfa.Text = ""
    DesClienteSurfa.Caption = ""
    ClientePelli.Text = ""
    DesClientePelli.Caption = ""
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    
    WEmpresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        ZZFecha = rstPedido!Fecha
        ZZCliente = ClienteSurfa.Text
        ZZFecEntrega = rstPedido!FecEntrega
        ZZHora = rstPedido!Hora
        ZZObservaciones = rstPedido!Observaciones
        ZZOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
        ZZMarca1 = rstPedido!Marca1
        ZZMarca2 = rstPedido!Marca2
        ZZMarca3 = rstPedido!Marca3
        ZZDestino = rstPedido!Destino
        ZZTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
        ZZVersion = "1"
        ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
        ZZVia = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
        
        rstPedido.Close
    
        Erase XEnvase
        Erase XEspecificaciones
        Erase XDatosMuestra
        Erase WPasaVector
        Erase Auxiliar
        
        Renglon = 0
        WRenglon = 0
    
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                        
                        WPasaVector(Renglon, 1) = rstPedido!Terminado
                        WPasaVector(Renglon, 2) = ""
                        WPasaVector(Renglon, 3) = Str$(rstPedido!Cantidad)
                        
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        XEspecificaciones(Renglon) = IIf(IsNull(rstPedido!Especificaciones), "0", rstPedido!Especificaciones)
                        
                        XDatosMuestra(Renglon, 1) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
                        XDatosMuestra(Renglon, 2) = IIf(IsNull(rstPedido!OrdenTrabajo), "", rstPedido!OrdenTrabajo)
                        XDatosMuestra(Renglon, 3) = IIf(IsNull(rstPedido!Referencia), "", rstPedido!Referencia)
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
        
        WRenglon = Renglon
        Renglon = 0
        
        For DA = 1 To WRenglon
        
            WLugar = DA
            Terminado = WPasaVector(WLugar, 1)
            
            WTipopro = "T"
            
            ZZDescripcion = ""
            spPrecios = "ConsultaPrecios " + "'" + ClientePelli.Text + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
    
                WPasaVector(WLugar, 2) = rstPrecios!Descripcion
                WPasaVector(WLugar, 4) = Str$(rstPrecios!Precio)
                ZZDescripcion = rstPrecios!Descripcion
                
                rstPrecios.Close
                
                If Trim(ZZDescripcion) = "" Then
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WPasaVector(WLugar, 2) = rstTerminado!Descripcion
                        rstTerminado.Close
                    End If
                End If
                
            End If
            
        Next DA
    
    End If
    
    Call Conecta_Empresa
    
    spCliente = "ConsultaCliente " + "'" + ClienteSurfa.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        ZZDesCliente = rstCliente!Razon
        ZZPago = Str(rstCliente!Pago1)
        rstCliente.Close
        
        spPago = "ConsultaPago " + "'" + ZZPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            ZZDesPago = rstPago!Nombre
            rstPago.Close
        End If
        
    End If
    
    ZZPedido = ""
    spPedido = "ListaPedidoNumero"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveLast
            ZZPedido = rstPedido!Pedido + 1
        End With
        rstPedido.Close
    End If
    
        
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 99
        
        WLugar = a
                
        Rem Articulo = WPasaVector(WLugar, 1)
        Articulo = "PE" + Mid$(WPasaVector(WLugar, 1), 3, 10)
        ArticuloPt = WPasaVector(WLugar, 1)
        NombreComercial = WPasaVector(WLugar, 2)
        Cantidad = WPasaVector(WLugar, 3)
        Precio = WPasaVector(WLugar, 4)
        
        If Val(Cantidad) <> 0 Then
        
        
        
            Rem dada
            Rem dada
            Rem dada
            Rem dada
            Rem dada
                    
        
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                rstTerminado.Close
            
                    Else
            
                XEmpresa = WEmpresa
                
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spTerminado = "ConsultaTerminado " + "'" + ArticuloPt + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                
                    ZZZZDescripcion = rstTerminado!Descripcion
                    ZZZZDescriEtiqueta = IIf(IsNull(rstTerminado!DescriEtiqueta), "", rstTerminado!DescriEtiqueta)
                    ZZZZLinea = rstTerminado!Linea
                    ZZZZUnidad = rstTerminado!Unidad
                    ZZZZInicial = Str$(rstTerminado!Inicial)
                    ZZZZEntradas = Str$(rstTerminado!Entradas)
                    ZZZZSalidas = Str$(rstTerminado!Salidas)
                    ZZZZMInimo = Str$(rstTerminado!MInimo)
                    ZZZZMinimo1 = IIf(IsNull(rstTerminado!Minimo1), "0", rstTerminado!Minimo1)
                    ZZZZFabrica = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
                    ZZZZFabricaII = IIf(IsNull(rstTerminado!FabricaII), "0", rstTerminado!FabricaII)
                    ZZZZFabricaIII = IIf(IsNull(rstTerminado!FabricaIII), "0", rstTerminado!FabricaIII)
                    ZZZZDeposito = IIf(IsNull(rstTerminado!Deposito), "", rstTerminado!Deposito)
                    ZZZZEnvase1 = rstTerminado!Envase1
                    ZZZZEnvase2 = rstTerminado!Envase2
                    ZZZZEnvase3 = rstTerminado!Envase3
                    ZZZZEnvase4 = rstTerminado!Envase4
                    ZZZZEnvase5 = rstTerminado!Envase5
                    ZZZZEnvase6 = rstTerminado!Envase6
                    ZZZZProceso = Str$(rstTerminado!Proceso)
                    ZZZZImpreadi = ""
                    ZZZZClase = ""
                    ZZZZSecundario = ""
                    ZZZZRiesgo = ""
                    ZZZZIntervencion = ""
                    ZZZZNaciones = ""
                    ZZZZEmbalaje = ""
                    ZZZZImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
                    ZZZZClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                    ZZZZSecundario = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
                    ZZZZRiesgo = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
                    ZZZZIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                    ZZZZNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                    ZZZZEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                    ZZZZControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    ZZZZSedronar = IIf(IsNull(rstTerminado!Sedronar), "0", rstTerminado!Sedronar)
                    ZZZZImpreVto = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
                    ZZZZMarca = IIf(IsNull(rstTerminado!Marca), "0", rstTerminado!Marca)
                    ZZZZObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
                    ZZZZTipoeti = IIf(IsNull(rstTerminado!Tipoeti), "", rstTerminado!Tipoeti)
                    ZZZZEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                    ZZZZPedido = Str$(rstTerminado!Pedido)
                    ZZZZConservacion = IIf(IsNull(rstTerminado!Conservacion), "", rstTerminado!Conservacion)
                    ZZZZConservacion = RTrim(Conservacion)
                    ZZZZConservacionII = IIf(IsNull(rstTerminado!ConservacionII), "", rstTerminado!ConservacionII)
                    ZZZZConservacionII = RTrim(ConservacionII)
                    ZZZZVida = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                    ZZZZSeguridad = IIf(IsNull(rstTerminado!seguridad), "", rstTerminado!seguridad)
                    
                    ZZZZVersion = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
                    ZZZZVersionI = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
                    ZZZZVersionII = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
                    
                    ZZZZFechaVersion = IIf(IsNull(rstTerminado!FechaVersion), "  /  /    ", rstTerminado!FechaVersion)
                    ZZZZFechaVersionI = IIf(IsNull(rstTerminado!FechaVersionI), "  /  /    ", rstTerminado!FechaVersionI)
                    ZZZZFechaVersionII = IIf(IsNull(rstTerminado!FechaVersionII), "  /  /    ", rstTerminado!FechaVersionII)
                    
                    ZZZZEstado = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                    ZZZZEstadoI = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                    ZZZZEstadoII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                    
                    ZZZZObserva = IIf(IsNull(rstTerminado!Observa), "", rstTerminado!Observa)
                    ZZZZObservaI = IIf(IsNull(rstTerminado!ObservaI), "", rstTerminado!ObservaI)
                    ZZZZObservaII = IIf(IsNull(rstTerminado!ObservaII), "", rstTerminado!ObservaII)
                    
                    ZZZZMetodo = IIf(IsNull(rstTerminado!Metodo), "", rstTerminado!Metodo)
                    ZZZZEfluentes = IIf(IsNull(rstTerminado!Efluentes), "", rstTerminado!Efluentes)
                    
                    ZZZZCaracteristicas = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
                    ZZZZCarga = IIf(IsNull(rstTerminado!Carga), "0", rstTerminado!Carga)
                    ZZZZEstadoProducto = IIf(IsNull(rstTerminado!EstadoProducto), "0", rstTerminado!EstadoProducto)
                    ZZZZListaProducto = IIf(IsNull(rstTerminado!ListaProducto), "0", rstTerminado!ListaProducto)
                    
                    ZZZZDescripcionIngles = IIf(IsNull(rstTerminado!Descripcioningles), "", rstTerminado!Descripcioningles)
                    ZZZZDescriEtiquetaIngles = IIf(IsNull(rstTerminado!DescriEtiquetaIngles), "", rstTerminado!DescriEtiquetaIngles)
                    ZZZZConservacionIngles = IIf(IsNull(rstTerminado!ConservacionIngles), "", rstTerminado!ConservacionIngles)
                    ZZZZConservacionIIIngles = IIf(IsNull(rstTerminado!ConservacionIIIngles), "", rstTerminado!ConservacionIIIngles)
                    
                    ZZZZResponsable = IIf(IsNull(rstTerminado!Responsable), "", rstTerminado!Responsable)
                    
                    ZZZZNaciones = Trim(ZZZZNaciones)
                    ZZZZImpreadi = Trim(ZZZZImpreadi)
                    ZZZZTipoeti = Trim(ZZZZTipoeti)
                    
                    rstTerminado.Close
                    
                End If
    
                Call Conecta_Empresa
    
    
                ZZZZLoteAutorizado = ""
                ZZZZPedido = ""
                ZZZZCosto = ""
                ZZZZFactor = ""
                ZZZZDate = Date$
                
                spTerminado = "ConsultaTerminado " + "'" + ArticuloPt + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount = 0 Then
                  
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO Terminado ("
                    ZSql = ZSql & "Codigo ,"
                    ZSql = ZSql & "Descripcion ,"
                    ZSql = ZSql & "DescriEtiqueta ,"
                    ZSql = ZSql & "Linea ,"
                    ZSql = ZSql & "Unidad ,"
                    ZSql = ZSql & "Inicial ,"
                    ZSql = ZSql & "Entradas ,"
                    ZSql = ZSql & "Salidas ,"
                    ZSql = ZSql & "Minimo ,"
                    ZSql = ZSql & "Minimo1 ,"
                    ZSql = ZSql & "Deposito ,"
                    ZSql = ZSql & "Pedido ,"
                    ZSql = ZSql & "Envase1 ,"
                    ZSql = ZSql & "Envase2 ,"
                    ZSql = ZSql & "Envase3 ,"
                    ZSql = ZSql & "Envase4 ,"
                    ZSql = ZSql & "Envase5 ,"
                    ZSql = ZSql & "Envase6 ,"
                    ZSql = ZSql & "Proceso ,"
                    ZSql = ZSql & "Costo ,"
                    ZSql = ZSql & "Factor ,"
                    ZSql = ZSql & "WDate ,"
                    ZSql = ZSql & "ImpreAdi ,"
                    ZSql = ZSql & "Clase ,"
                    ZSql = ZSql & "Secundario ,"
                    ZSql = ZSql & "Riesgo ,"
                    ZSql = ZSql & "Intervencion ,"
                    ZSql = ZSql & "Naciones ,"
                    ZSql = ZSql & "Embalaje ,"
                    ZSql = ZSql & "Controla ,"
                    ZSql = ZSql & "Sedronar ,"
                    ZSql = ZSql & "ImpreVto ,"
                    ZSql = ZSql & "Marca ,"
                    ZSql = ZSql & "Observaciones ,"
                    ZSql = ZSql & "TipoEti ,"
                    ZSql = ZSql & "Escrito ,"
                    ZSql = ZSql & "Fabrica ,"
                    ZSql = ZSql & "FabricaII ,"
                    ZSql = ZSql & "FabricaIII ,"
                    ZSql = ZSql & "LoteAutorizado ,"
                    ZSql = ZSql & "Conservacion ,"
                    ZSql = ZSql & "ConservacionII ,"
                    ZSql = ZSql & "Vida ,"
                    ZSql = ZSql & "Seguridad ,"
                    ZSql = ZSql & "Version ,"
                    ZSql = ZSql & "VersionI ,"
                    ZSql = ZSql & "VersionII ,"
                    ZSql = ZSql & "FechaVersion ,"
                    ZSql = ZSql & "FechaVersionI ,"
                    ZSql = ZSql & "FechaVersionII ,"
                    ZSql = ZSql & "Estado ,"
                    ZSql = ZSql & "EstadoI ,"
                    ZSql = ZSql & "EstadoII ,"
                    ZSql = ZSql & "Observa ,"
                    ZSql = ZSql & "ObservaI ,"
                    ZSql = ZSql & "ObservaII ,"
                    ZSql = ZSql & "DescripcionIngles ,"
                    ZSql = ZSql & "DescriEtiquetaIngles ,"
                    ZSql = ZSql & "ConservacionIngles ,"
                    ZSql = ZSql & "ConservacionIIIngles ,"
                    ZSql = ZSql & "Metodo ,"
                    ZSql = ZSql & "Efluentes )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + Articulo + "',"
                    ZSql = ZSql & "'" + ZZZZDescripcion + "',"
                    ZSql = ZSql & "'" + ZZZZDescriEtiqueta + "',"
                    ZSql = ZSql & "'" + ZZZZLinea + "',"
                    ZSql = ZSql & "'" + ZZZZUnidad + "',"
                    ZSql = ZSql & "'" + ZZZZInicial + "',"
                    ZSql = ZSql & "'" + ZZZZEntradas + "',"
                    ZSql = ZSql & "'" + ZZZZSalidas + "',"
                    ZSql = ZSql & "'" + ZZZZMInimo + "',"
                    ZSql = ZSql & "'" + ZZZZMinimo1 + "',"
                    ZSql = ZSql & "'" + ZZZZDeposito + "',"
                    ZSql = ZSql & "'" + ZZZZPedido + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase1 + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase2 + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase3 + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase4 + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase5 + "',"
                    ZSql = ZSql & "'" + ZZZZEnvase6 + "',"
                    ZSql = ZSql & "'" + ZZZZProceso + "',"
                    ZSql = ZSql & "'" + ZZZZCosto + "',"
                    ZSql = ZSql & "'" + ZZZZFactor + "',"
                    ZSql = ZSql & "'" + ZZZZDate + "',"
                    ZSql = ZSql & "'" + ZZZZImpreadi + "',"
                    ZSql = ZSql & "'" + ZZZZClase + "',"
                    ZSql = ZSql & "'" + ZZZZSecundario + "',"
                    ZSql = ZSql & "'" + ZZZZRiesgo + "',"
                    ZSql = ZSql & "'" + ZZZZIntervencion + "',"
                    ZSql = ZSql & "'" + ZZZZNaciones + "',"
                    ZSql = ZSql & "'" + ZZZZEmbalaje + "',"
                    ZSql = ZSql & "'" + ZZZZControla + "',"
                    ZSql = ZSql & "'" + ZZZZSedronar + "',"
                    ZSql = ZSql & "'" + ZZZZImpreVto + "',"
                    ZSql = ZSql & "'" + ZZZZMarca + "',"
                    ZSql = ZSql & "'" + ZZZZObservaciones + "',"
                    ZSql = ZSql & "'" + ZZZZTipoeti + "',"
                    ZSql = ZSql & "'" + ZZZZEscrito + "',"
                    ZSql = ZSql & "'" + ZZZZFabrica + "',"
                    ZSql = ZSql & "'" + ZZZZFabricaII + "',"
                    ZSql = ZSql & "'" + ZZZZFabricaIII + "',"
                    ZSql = ZSql & "'" + ZZZZLoteAutorizado + "',"
                    ZSql = ZSql & "'" + ZZZZConservacion + "',"
                    ZSql = ZSql & "'" + ZZZZConservacionII + "',"
                    ZSql = ZSql & "'" + ZZZZVida + "',"
                    ZSql = ZSql & "'" + ZZZZSeguridad + "',"
                    ZSql = ZSql & "'" + ZZZZVersion + "',"
                    ZSql = ZSql & "'" + ZZZZVersionI + "',"
                    ZSql = ZSql & "'" + ZZZZVersionII + "',"
                    ZSql = ZSql & "'" + ZZZZFechaVersion + "',"
                    ZSql = ZSql & "'" + ZZZZFechaVersionI + "',"
                    ZSql = ZSql & "'" + ZZZZFechaVersionII + "',"
                    ZSql = ZSql & "'" + ZZZZEstado + "',"
                    ZSql = ZSql & "'" + ZZZZEstadoI + "',"
                    ZSql = ZSql & "'" + ZZZZEstadoII + "',"
                    ZSql = ZSql & "'" + ZZZZObserva + "',"
                    ZSql = ZSql & "'" + ZZZZObservaI + "',"
                    ZSql = ZSql & "'" + ZZZZObservaII + "',"
                    ZSql = ZSql & "'" + ZZZZDescripcionIngles + "',"
                    ZSql = ZSql & "'" + ZZZZDescriEtiquetaIngles + "',"
                    ZSql = ZSql & "'" + ZZZZConservacionIngles + "',"
                    ZSql = ZSql & "'" + ZZZZConservacionIIIngles + "',"
                    ZSql = ZSql & "'" + ZZZZMetodo + "',"
                    ZSql = ZSql & "'" + ZZZZEfluentes + "')"
        
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
              
                          Else
                          
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Terminado SET "
                    ZSql = ZSql & "Codigo = " + "'" + Articulo + "',"
                    ZSql = ZSql & "Descripcion = " + "'" + ZZZZDescripcion + "',"
                    ZSql = ZSql & "DescriEtiqueta = " + "'" + ZZZZDescriEtiqueta + "',"
                    ZSql = ZSql & "Linea = " + "'" + ZZZZLinea + "',"
                    ZSql = ZSql & "Unidad = " + "'" + ZZZZUnidad + "',"
                    ZSql = ZSql & "Inicial = " + "'" + ZZZZInicial + "',"
                    ZSql = ZSql & "Entradas = " + "'" + ZZZZEntradas + "',"
                    ZSql = ZSql & "Salidas = " + "'" + ZZZZSalidas + "',"
                    ZSql = ZSql & "Minimo = " + "'" + ZZZZMInimo + "',"
                    ZSql = ZSql & "Minimo1 = " + "'" + ZZZZMinimo1 + "',"
                    ZSql = ZSql & "Deposito = " + "'" + ZZZZDeposito + "',"
                    ZSql = ZSql & "Pedido = " + "'" + ZZZZPedido + "',"
                    ZSql = ZSql & "Envase1 = " + "'" + ZZZZEnvase1 + "',"
                    ZSql = ZSql & "Envase2 = " + "'" + ZZZZEnvase2 + "',"
                    ZSql = ZSql & "Envase3 = " + "'" + ZZZZEnvase3 + "',"
                    ZSql = ZSql & "Envase4 = " + "'" + ZZZZEnvase4 + "',"
                    ZSql = ZSql & "Envase5 = " + "'" + ZZZZEnvase5 + "',"
                    ZSql = ZSql & "Envase6 = " + "'" + ZZZZEnvase6 + "',"
                    ZSql = ZSql & "Proceso = " + "'" + ZZZZProceso + "',"
                    ZSql = ZSql & "Costo = " + "'" + ZZZZCosto + "',"
                    ZSql = ZSql & "Factor = " + "'" + ZZZZFactor + "',"
                    ZSql = ZSql & "WDate = " + "'" + ZZZZDate + "',"
                    ZSql = ZSql & "ImpreAdi = " + "'" + ZZZZImpreadi + "',"
                    ZSql = ZSql & "Clase = " + "'" + ZZZZClase + "',"
                    ZSql = ZSql & "Secundario = " + "'" + ZZZZSecundario + "',"
                    ZSql = ZSql & "Riesgo = " + "'" + ZZZZRiesgo + "',"
                    ZSql = ZSql & "Intervencion = " + "'" + ZZZZIntervencion + "',"
                    ZSql = ZSql & "Naciones = " + "'" + ZZZZNaciones + "',"
                    ZSql = ZSql & "Embalaje = " + "'" + ZZZZEmbalaje + "',"
                    ZSql = ZSql & "Controla = " + "'" + ZZZZControla + "',"
                    ZSql = ZSql & "Sedronar = " + "'" + ZZZZSedronar + "',"
                    ZSql = ZSql & "ImpreVto = " + "'" + ZZZZImpreVto + "',"
                    ZSql = ZSql & "Marca = " + "'" + ZZZZMarca + "',"
                    ZSql = ZSql & "Observaciones = " + "'" + ZZZZObservaciones + "',"
                    ZSql = ZSql & "TipoEti = " + "'" + ZZZZTipoeti + "',"
                    ZSql = ZSql & "Escrito = " + "'" + ZZZZEscrito + "',"
                    ZSql = ZSql & "Fabrica = " + "'" + ZZZZFabrica + "',"
                    ZSql = ZSql & "FabricaII = " + "'" + ZZZZFabricaII + "',"
                    ZSql = ZSql & "FabricaIII = " + "'" + ZZZZFabricaIII + "',"
                    ZSql = ZSql & "LoteAutorizado = " + "'" + ZZZZLoteAutorizado + "',"
                    ZSql = ZSql & "Conservacion = " + "'" + ZZZZConservacion + "',"
                    ZSql = ZSql & "ConservacionII = " + "'" + ZZZZConservacionII + "',"
                    ZSql = ZSql & "Vida = " + "'" + ZZZZVida + "',"
                    ZSql = ZSql & "Seguridad = " + "'" + ZZZZSeguridad + "',"
                    ZSql = ZSql & "Version = " + "'" + ZZZZVersion + "',"
                    ZSql = ZSql & "VersionI = " + "'" + ZZZZVersionI + "',"
                    ZSql = ZSql & "VersionII = " + "'" + ZZZZVersionII + "',"
                    ZSql = ZSql & "FechaVersion = " + "'" + ZZZZFechaVersion + "',"
                    ZSql = ZSql & "FechaVersionI = " + "'" + ZZZZFechaVersionI + "',"
                    ZSql = ZSql & "FechaVersionII = " + "'" + ZZZZFechaVersionII + "',"
                    ZSql = ZSql & "Estado = " + "'" + ZZZZEstado + "',"
                    ZSql = ZSql & "EstadoI = " + "'" + ZZZZEstadoI + "',"
                    ZSql = ZSql & "EstadoII = " + "'" + ZZZZEstadoII + "',"
                    ZSql = ZSql & "Observa = " + "'" + ZZZZObserva + "',"
                    ZSql = ZSql & "ObservaI = " + "'" + ZZZZObservaI + "',"
                    ZSql = ZSql & "ObservaII = " + "'" + ZZZZObservaII + "',"
                    ZSql = ZSql & "DescripcionIngles = " + "'" + ZZZZDescripcionIngles + "',"
                    ZSql = ZSql & "DescriEtiquetaIngles = " + "'" + ZZZZDescriEtiquetaIngles + "',"
                    ZSql = ZSql & "ConservacionIngles = " + "'" + zzzzzConservacionIngles + "',"
                    ZSql = ZSql & "ConservacionIIIngles = " + "'" + ZZZZConservacionIIIngles + "',"
                    ZSql = ZSql & "Metodo = " + "'" + ZZZZMetodo + "',"
                    ZSql = ZSql & "Efluentes = " + "'" + ZZZZEfluentes + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + Articulo + "'"
                  
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
              
                End If
              
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "Responsable = " + "'" + ZZZZResponsable + "',"
                ZSql = ZSql & "DescriOnu = " + "'" + ZZZZCaracteristicas + "',"
                ZSql = ZSql & "Carga = " + "'" + ZZZZCarga + "',"
                ZSql = ZSql & "EstadoProducto = " + "'" + ZZZZEstadoProducto + "',"
                ZSql = ZSql & "ListaProducto = " + "'" + ZZZZListaProducto + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + Articulo + "'"
                  
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
              
              
              
              
              
                Rem verifica la alta en todas las empresas
              
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
                CargaEmpresa(6, 1) = "0010"
                CargaEmpresa(6, 2) = "Empresa10"
                CargaEmpresa(7, 1) = "0011"
                CargaEmpresa(7, 2) = "Empresa11"
                  
                For Cicla = 1 To 7
              
                    If CargaEmpresa(Cicla, 1) <> "" Then
              
                        WEmpresa = CargaEmpresa(Cicla, 1)
                        txtOdbc = CargaEmpresa(Cicla, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                      
                      
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount = 0 Then
              
                            ZSql = ""
                            ZSql = ZSql & "INSERT INTO Terminado ("
                            ZSql = ZSql & "Codigo ,"
                            ZSql = ZSql & "Descripcion ,"
                            ZSql = ZSql & "DescriEtiqueta ,"
                            ZSql = ZSql & "Linea ,"
                            ZSql = ZSql & "Unidad ,"
                            ZSql = ZSql & "Inicial ,"
                            ZSql = ZSql & "Entradas ,"
                            ZSql = ZSql & "Salidas ,"
                            ZSql = ZSql & "Minimo ,"
                            ZSql = ZSql & "Minimo1 ,"
                            ZSql = ZSql & "Deposito ,"
                            ZSql = ZSql & "Pedido ,"
                            ZSql = ZSql & "Envase1 ,"
                            ZSql = ZSql & "Envase2 ,"
                            ZSql = ZSql & "Envase3 ,"
                            ZSql = ZSql & "Envase4 ,"
                            ZSql = ZSql & "Envase5 ,"
                            ZSql = ZSql & "Envase6 ,"
                            ZSql = ZSql & "Proceso ,"
                            ZSql = ZSql & "Costo ,"
                            ZSql = ZSql & "Factor ,"
                            ZSql = ZSql & "WDate ,"
                            ZSql = ZSql & "ImpreAdi ,"
                            ZSql = ZSql & "Clase ,"
                            ZSql = ZSql & "Secundario ,"
                            ZSql = ZSql & "Riesgo ,"
                            ZSql = ZSql & "Intervencion ,"
                            ZSql = ZSql & "Naciones ,"
                            ZSql = ZSql & "Embalaje ,"
                            ZSql = ZSql & "Controla ,"
                            ZSql = ZSql & "Sedronar ,"
                            ZSql = ZSql & "ImpreVto ,"
                            ZSql = ZSql & "Marca ,"
                            ZSql = ZSql & "Observaciones ,"
                            ZSql = ZSql & "TipoEti ,"
                            ZSql = ZSql & "Escrito ,"
                            ZSql = ZSql & "Fabrica ,"
                            ZSql = ZSql & "FabricaII ,"
                            ZSql = ZSql & "FabricaIII ,"
                            ZSql = ZSql & "LoteAutorizado ,"
                            ZSql = ZSql & "Conservacion ,"
                            ZSql = ZSql & "ConservacionII ,"
                            ZSql = ZSql & "Vida ,"
                            ZSql = ZSql & "Seguridad ,"
                            ZSql = ZSql & "Version ,"
                            ZSql = ZSql & "VersionI ,"
                            ZSql = ZSql & "VersionII ,"
                            ZSql = ZSql & "FechaVersion ,"
                            ZSql = ZSql & "FechaVersionI ,"
                            ZSql = ZSql & "FechaVersionII ,"
                            ZSql = ZSql & "Estado ,"
                            ZSql = ZSql & "EstadoI ,"
                            ZSql = ZSql & "EstadoII ,"
                            ZSql = ZSql & "Observa ,"
                            ZSql = ZSql & "ObservaI ,"
                            ZSql = ZSql & "ObservaII ,"
                            ZSql = ZSql & "DescripcionIngles ,"
                            ZSql = ZSql & "DescriEtiquetaIngles ,"
                            ZSql = ZSql & "ConservacionIngles ,"
                            ZSql = ZSql & "ConservacionIIIngles ,"
                            ZSql = ZSql & "Metodo ,"
                            ZSql = ZSql & "Efluentes )"
                            ZSql = ZSql & "Values ("
                            ZSql = ZSql & "'" + Articulo + "',"
                            ZSql = ZSql & "'" + ZZZZDescripcion + "',"
                            ZSql = ZSql & "'" + ZZZZDescriEtiqueta + "',"
                            ZSql = ZSql & "'" + ZZZZLinea + "',"
                            ZSql = ZSql & "'" + ZZZZUnidad + "',"
                            ZSql = ZSql & "'" + ZZZZInicial + "',"
                            ZSql = ZSql & "'" + ZZZZEntradas + "',"
                            ZSql = ZSql & "'" + ZZZZSalidas + "',"
                            ZSql = ZSql & "'" + ZZZZMInimo + "',"
                            ZSql = ZSql & "'" + ZZZZMinimo1 + "',"
                            ZSql = ZSql & "'" + ZZZZDeposito + "',"
                            ZSql = ZSql & "'" + ZZZZPedido + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase1 + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase2 + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase3 + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase4 + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase5 + "',"
                            ZSql = ZSql & "'" + ZZZZEnvase6 + "',"
                            ZSql = ZSql & "'" + ZZZZProceso + "',"
                            ZSql = ZSql & "'" + ZZZZCosto + "',"
                            ZSql = ZSql & "'" + ZZZZFactor + "',"
                            ZSql = ZSql & "'" + ZZZZDate + "',"
                            ZSql = ZSql & "'" + ZZZZImpreadi + "',"
                            ZSql = ZSql & "'" + ZZZZClase + "',"
                            ZSql = ZSql & "'" + ZZZZSecundario + "',"
                            ZSql = ZSql & "'" + ZZZZRiesgo + "',"
                            ZSql = ZSql & "'" + ZZZZIntervencion + "',"
                            ZSql = ZSql & "'" + ZZZZNaciones + "',"
                            ZSql = ZSql & "'" + ZZZZEmbalaje + "',"
                            ZSql = ZSql & "'" + ZZZZControla + "',"
                            ZSql = ZSql & "'" + ZZZZSedronar + "',"
                            ZSql = ZSql & "'" + ZZZZImpreVto + "',"
                            ZSql = ZSql & "'" + ZZZZMarca + "',"
                            ZSql = ZSql & "'" + ZZZZObservaciones + "',"
                            ZSql = ZSql & "'" + ZZZZTipoeti + "',"
                            ZSql = ZSql & "'" + ZZZZEscrito + "',"
                            ZSql = ZSql & "'" + ZZZZFabrica + "',"
                            ZSql = ZSql & "'" + ZZZZFabricaII + "',"
                            ZSql = ZSql & "'" + ZZZZFabricaIII + "',"
                            ZSql = ZSql & "'" + ZZZZLoteAutorizado + "',"
                            ZSql = ZSql & "'" + ZZZZConservacion + "',"
                            ZSql = ZSql & "'" + ZZZZConservacionII + "',"
                            ZSql = ZSql & "'" + ZZZZVida + "',"
                            ZSql = ZSql & "'" + ZZZZSeguridad + "',"
                            ZSql = ZSql & "'" + ZZZZVersion + "',"
                            ZSql = ZSql & "'" + ZZZZVersionI + "',"
                            ZSql = ZSql & "'" + ZZZZVersionII + "',"
                            ZSql = ZSql & "'" + ZZZZFechaVersion + "',"
                            ZSql = ZSql & "'" + ZZZZFechaVersionI + "',"
                            ZSql = ZSql & "'" + ZZZZFechaVersionII + "',"
                            ZSql = ZSql & "'" + ZZZZEstado + "',"
                            ZSql = ZSql & "'" + ZZZZEstadoI + "',"
                            ZSql = ZSql & "'" + ZZZZEstadoII + "',"
                            ZSql = ZSql & "'" + ZZZZObserva + "',"
                            ZSql = ZSql & "'" + ZZZZObservaI + "',"
                            ZSql = ZSql & "'" + ZZZZObservaII + "',"
                            ZSql = ZSql & "'" + ZZZZDescripcionIngles + "',"
                            ZSql = ZSql & "'" + ZZZZDescriEtiquetaIngles + "',"
                            ZSql = ZSql & "'" + ZZZZConservacionIngles + "',"
                            ZSql = ZSql & "'" + ZZZZConservacionIIIngles + "',"
                            ZSql = ZSql & "'" + ZZZZMetodo + "',"
                            ZSql = ZSql & "'" + ZZZZEfluentes + "')"
        
                            spTerminado = ZSql
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
              
                                  Else
                          
                            ZSql = ""
                            ZSql = ZSql & "UPDATE Terminado SET "
                            ZSql = ZSql & "Codigo = " + "'" + zzzzCodigo + "',"
                            ZSql = ZSql & "Descripcion = " + "'" + ZZZZDescripcion + "',"
                            ZSql = ZSql & "DescriEtiqueta = " + "'" + ZZZZDescriEtiqueta + "',"
                            ZSql = ZSql & "Linea = " + "'" + ZZZZLinea + "',"
                            ZSql = ZSql & "Unidad = " + "'" + ZZZZUnidad + "',"
                            ZSql = ZSql & "Minimo = " + "'" + ZZZZMInimo + "',"
                            ZSql = ZSql & "Minimo1 = " + "'" + ZZZZMinimo1 + "',"
                            ZSql = ZSql & "Deposito = " + "'" + ZZZZDeposito + "',"
                            ZSql = ZSql & "Envase1 = " + "'" + ZZZZEnvase1 + "',"
                            ZSql = ZSql & "Envase2 = " + "'" + ZZZZEnvase2 + "',"
                            ZSql = ZSql & "Envase3 = " + "'" + ZZZZEnvase3 + "',"
                            ZSql = ZSql & "Envase4 = " + "'" + ZZZZEnvase4 + "',"
                            ZSql = ZSql & "Envase5 = " + "'" + ZZZZEnvase5 + "',"
                            ZSql = ZSql & "Envase6 = " + "'" + ZZZZEnvase6 + "',"
                            ZSql = ZSql & "Costo = " + "'" + ZZZZCosto + "',"
                            ZSql = ZSql & "Factor = " + "'" + ZZZZFactor + "',"
                            ZSql = ZSql & "WDate = " + "'" + ZZZZDate + "',"
                            ZSql = ZSql & "ImpreAdi = " + "'" + ZZZZImpreadi + "',"
                            ZSql = ZSql & "Clase = " + "'" + ZZZZClase + "',"
                            ZSql = ZSql & "Secundario = " + "'" + ZZZZSecundario + "',"
                            ZSql = ZSql & "Riesgo = " + "'" + ZZZZRiesgo + "',"
                            ZSql = ZSql & "Intervencion = " + "'" + ZZZZIntervencion + "',"
                            ZSql = ZSql & "Naciones = " + "'" + ZZZZNaciones + "',"
                            ZSql = ZSql & "Embalaje = " + "'" + ZZZZEmbalaje + "',"
                            ZSql = ZSql & "Controla = " + "'" + ZZZZControla + "',"
                            ZSql = ZSql & "Sedronar = " + "'" + ZZZZSedronar + "',"
                            ZSql = ZSql & "ImpreVto = " + "'" + ZZZZImpreVto + "',"
                            ZSql = ZSql & "Marca = " + "'" + ZZZZMarca + "',"
                            ZSql = ZSql & "Observaciones = " + "'" + ZZZZObservaciones + "',"
                            ZSql = ZSql & "TipoEti = " + "'" + ZZZZTipoeti + "',"
                            ZSql = ZSql & "Escrito = " + "'" + ZZZZEscrito + "',"
                            ZSql = ZSql & "Fabrica = " + "'" + ZZZZFabrica + "',"
                            ZSql = ZSql & "FabricaII = " + "'" + ZZZZFabricaII + "',"
                            ZSql = ZSql & "FabricaIII = " + "'" + ZZZZFabricaIII + "',"
                            ZSql = ZSql & "LoteAutorizado = " + "'" + ZZZZLoteAutorizado + "',"
                            ZSql = ZSql & "Conservacion = " + "'" + ZZZZConservacion + "',"
                            ZSql = ZSql & "ConservacionII = " + "'" + ZZZZConservacionII + "',"
                            ZSql = ZSql & "Vida = " + "'" + ZZZZVida + "',"
                            ZSql = ZSql & "Seguridad = " + "'" + ZZZZSeguridad + "',"
                            ZSql = ZSql & "Version = " + "'" + ZZZZVersion + "',"
                            ZSql = ZSql & "VersionI = " + "'" + ZZZZVersionI + "',"
                            ZSql = ZSql & "VersionII = " + "'" + ZZZZVersionII + "',"
                            ZSql = ZSql & "FechaVersion = " + "'" + ZZZZFechaVersion + "',"
                            ZSql = ZSql & "FechaVersionI = " + "'" + ZZZZFechaVersionI + "',"
                            ZSql = ZSql & "FechaVersionII = " + "'" + ZZZZFechaVersionII + "',"
                            ZSql = ZSql & "Estado = " + "'" + ZZZZEstado + "',"
                            ZSql = ZSql & "EstadoI = " + "'" + ZZZZEstadoI + "',"
                            ZSql = ZSql & "EstadoII = " + "'" + ZZZZEstadoII + "',"
                            ZSql = ZSql & "Observa = " + "'" + ZZZZObserva + "',"
                            ZSql = ZSql & "ObservaI = " + "'" + ZZZZObservaI + "',"
                            ZSql = ZSql & "ObservaII = " + "'" + ZZZZObservaII + "',"
                            ZSql = ZSql & "DescripcionIngles = " + "'" + ZZZZDescripcionIngles + "',"
                            ZSql = ZSql & "DescriEtiquetaIngles = " + "'" + ZZZZDescriEtiquetaIngles + "',"
                            ZSql = ZSql & "ConservacionIngles = " + "'" + ZZZZConservacionIngles + "',"
                            ZSql = ZSql & "ConservacionIIIngles = " + "'" + ZZZZConservacionIIIngles + "',"
                            ZSql = ZSql & "Metodo = " + "'" + ZZZZMetodo + "',"
                            ZSql = ZSql & "Efluentes = " + "'" + ZZZZEfluentes + "'"
                            ZSql = ZSql & " Where Codigo = " + "'" + Articulo + "'"
                          
                            spTerminado = ZSql
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                          
                        End If
                      
                        ZSql = ""
                        ZSql = ZSql & "UPDATE Terminado SET "
                        ZSql = ZSql & "Responsable = " + "'" + ZZZZResponsable + "',"
                        ZSql = ZSql & "DescriOnu = " + "'" + ZZZZCaracteristicas + "',"
                        ZSql = ZSql & "Carga = " + "'" + ZZZZCarga + "',"
                        ZSql = ZSql & "EstadoProducto = " + "'" + ZZZZEstadoProducto + "',"
                        ZSql = ZSql & "ListaProducto = " + "'" + ZZZZListaProducto + "'"
                        ZSql = ZSql & " Where Codigo = " + "'" + Articulo + "'"
                          
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                      
                Next Cicla
                
                Call Conecta_Empresa
            
            End If
        
                    
                    
                    
                            
            
            XEmpresa = WEmpresa
            
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZZZZPrecioClave = ClientePelli.Text + ArticuloPt
            ZZZZPrecioPrecio = ""
            ZZZZPrecioDescripcion = ""
            ZZZZPrecioFecha = ""
            ZZZZPrecioPago = ""
            ZZZZPrecioEstado = ""
            
            spPrecios = "ConsultaPrecios " + "'" + ZZZZPrecioClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                ZZZZPrecioPrecio = Str$(rstPrecios!Precio)
                ZZZZPrecioDescripcion = rstPrecios!Descripcion
                ZZZZPrecioFecha = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
                ZZZZPrecioPago = IIf(IsNull(rstPrecios!Pago), "0", rstPrecios!Pago)
                ZZZZPrecioEstado = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
                    
                rstPrecios.Close
        
            End If
                
            Call Conecta_Empresa
                        


            ZZZZPrecioClave = ClienteSurfa.Text + Articulo
            ZZZZPrecioFecha1 = ""
            ZZZZPrecioFecha2 = ""
            ZZZZPrecioFecha3 = ""
            ZZZZPrecioFecha4 = ""
            ZZZZPrecioFecha5 = ""
            ZZZZPrecioFecha6 = ""
            ZZZZPrecioFactura1 = ""
            ZZZZPrecioFactura2 = ""
            ZZZZPrecioFactura3 = ""
            ZZZZPrecioFactura4 = ""
            ZZZZPrecioFactura5 = ""
            ZZZZPrecioPrecio1 = ""
            ZZZZPrecioPrecio2 = ""
            ZZZZPrecioPrecio3 = ""
            ZZZZPrecioPrecio4 = ""
            ZZZZPrecioPrecio5 = ""
    
            spPrecios = "ConsultaPrecios " + "'" + ZZZZPrecioClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                rstPrecios.Close
                ZSql = ""
                ZSql = ZSql & "UPDATE Precios SET "
                ZSql = ZSql & "Descripcion = " + "'" + ZZZZPrecioDescripcion + "',"
                ZSql = ZSql & "Precio = " + "'" + ZZZZPrecioPrecio + "'"
                ZSql = ZSql & " Where Clave = " + "'" + ZZZZPrecioClave + "'"
                spPrecios = ZSql
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + ZZZZPrecioClave + "','" + ClienteSurfa.Text + "','" + Articulo + "','" + ZZZZPrecioPrecio + "','" _
                         + ZZZZPrecioDescripcion + "','" _
                         + ZZZZPrecioFecha1 + "','" + ZZZZPrecioFactura1 + "','" + ZZZZPrecioPrecio1 + "','" + ZZZZPrecioCantidad1 + "','" _
                         + ZZZZPrecioFecha2 + "','" + ZZZZPrecioFactura2 + "','" + ZZZZPrecioPrecio2 + "','" + ZZZZPrecioCantidad2 + "','" _
                         + ZZZZPrecioFecha3 + "','" + ZZZZPrecioFactura3 + "','" + ZZZZPrecioPrecio3 + "','" + ZZZZPrecioCantidad3 + "','" _
                         + ZZZZPrecioFecha4 + "','" + ZZZZPrecioFactura4 + "','" + ZZZZPrecioPrecio4 + "','" + ZZZZPrecioCantidad4 + "','" _
                         + ZZZZPrecioFecha5 + "','" + ZZZZPrecioFactura5 + "','" + ZZZZPrecioPrecio5 + "','" + ZZZZPrecioCantidad5 + "','" _
                         + Date$ + "','" + ZZZZPrecioFecha + "','" + ZZZZPrecioPago + "'"
                Set rstPrecios = db.OpenRecordset("AltaPrecios1 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
    
        
            ZSql = ""
            ZSql = ZSql & "UPDATE Precios SET "
            ZSql = ZSql & "Estado = " + "'" + ZZZZPrecioEstado + "'"
            ZSql = ZSql & " Where Clave = " + "'" + ZZZZPrecioClave + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        
          
          
          
        
        
        
            Renglon = Renglon + 1
            WRenglon = WRenglon + 1
                
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                    
            Auxi1 = ZZPedido
            Call Ceros(Auxi1, 6)
                
            WPedido = Str$(ZZPedido)
            WRenglon = Str$(Renglon)
            WCliente = ClienteSurfa.Text
            WFecha = ZZFecha
            WFecEntrega = ZZFecEntrega
            WHora = ZZHora
            WObservaciones = ZZObservaciones
            WOrdenCpa = ZZOrdenCpa
            WTerminado = Articulo
            WCantidad = Cantidad
            WEnvase1 = XEnvase(WLugar, 1)
            WCanti1 = XEnvase(WLugar, 2)
            WEnvase2 = XEnvase(WLugar, 3)
            WCanti2 = XEnvase(WLugar, 4)
            WEnvase3 = XEnvase(WLugar, 5)
            WCanti3 = XEnvase(WLugar, 6)
            WEspecificaciones = XEspecificaciones(WLugar)
            WOrdenTrabajo = XDatosMuestra(WLugar, 2)
            WReferencia = XDatosMuestra(WLugar, 3)
            WEnvase4 = ""
            WCanti4 = ""
            WFechaord = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
            WOrdFecEntrega = Right$(ZZFecEntrega, 4) + Mid$(ZZFecEntrega, 4, 2) + Left$(ZZFecEntrega, 2)
            WPrecio = Precio
            WWLinea = "1"
            WFacturado = ""
            WImporte = ""
            WClave = Auxi1 + Auxi
            WMarca1 = ZZMarca1
            WMarca2 = ZZMarca2
            WMarca3 = ZZMarca3
            WDestino = ZZDestino
            WAutorizo = "S"
            WImpresion = "S"
            WTipoPed = ZZTipoped
            WCantidad1 = ""
            WCantidad2 = ""
            WLote1 = "0"
            WLote2 = "0"
            Wlote3 = "0"
            WLote4 = "0"
            WLote5 = "0"
            WCantiLote1 = "0"
            WCantiLote2 = "0"
            WCantiLote3 = "0"
            WCantiLote4 = "0"
            WCantiLote5 = "0"
            WEnv1 = "0"
            WEnv2 = "0"
            WEnv3 = "0"
            WEnv4 = "0"
            WEnv5 = "0"
            WCantiEnv1 = "0"
            WCantiEnv2 = "0"
            WCantiEnv3 = "0"
            WCantiEnv4 = "0"
            WCantiEnv5 = "0"
            WVersion = "1"
            WTipopro = "T"
            WArti = "  -   -   "
            WVia = ZZVia
            
            XParam = "'" + WClave + "','" _
                         + WPedido + "','" + WRenglon + "','" + WCliente + "','" + WFecha + "','" _
                         + WFecEntrega + "','" + WHora + "','" + WObservaciones + "','" + WTerminado + "','" _
                         + WCantidad + "','" + WEnvase1 + "','" + WCanti1 + "','" _
                         + WEnvase2 + "','" + WCanti2 + "','" _
                         + WEnvase3 + "','" + WCanti3 + "','" _
                         + WEnvase4 + "','" + WCanti4 + "','" _
                         + WFechaord + "','" + WPrecio + "','" + WWLinea + "','" + WFacturado + "','" _
                         + WImporte + "','" + WMarca1 + "','" _
                         + WMarca2 + "','" + WMarca3 + "','" + WDestino + "','" _
                         + WAutorizo + "','" + WImpresion + "','" + WTipoPed + "','" _
                         + WCantidad1 + "','" + WCantidad2 + "','" _
                         + WLote1 + "','" + WCantiLote1 + "','" + WLote2 + "','" + WCantiLote2 + "','" _
                         + Wlote3 + "','" + WCantiLote3 + "','" + WLote1 + "','" + WCantiLote4 + "','" _
                         + WLote5 + "','" + WCantiLote5 + "','" _
                         + WEnv1 + "','" + WCantiEnv1 + "','" + WEnv2 + "','" + WCantiEnv2 + "','" _
                         + WEnv3 + "','" + WCantiEnv3 + "','" + WEnv4 + "','" + WCantiEnv4 + "','" _
                         + WEnv5 + "','" + WCantiEnv5 + "','" _
                         + WVersion + "','" _
                         + WOrdFecEntrega + "','" _
                         + WOrdenCpa + "','" _
                         + WTipopro + "','" _
                         + WArti + "'"

            spPedido = "AltaPedido " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next a
    
    Call Cierre_Click
        
End Sub
        
Private Sub Pedido_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Fecha.Text = rstPedido!Fecha
            ClienteSurfa.SetFocus
            rstPedido.Close
        End If
        
        Call Conecta_Empresa
        
    End If

    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub ClienteSurfa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        ClienteSurfa.Text = UCase(ClienteSurfa.Text)
        
        spCliente = "ConsultaCliente " + "'" + ClienteSurfa.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesClienteSurfa.Caption = rstCliente!Razon
            ClientePelli.SetFocus
            rstCliente.Close
        End If
        
    End If
    
End Sub

Private Sub ClientePelli_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ClientePelli.Text = UCase(ClientePelli.Text)
        
        spCliente = "ConsultaCliente " + "'" + ClientePelli.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesClientePelli.Caption = rstCliente!Razon
            Pedido.SetFocus
            rstCliente.Close
        End If
        
        Call Conecta_Empresa
        
    End If
    
End Sub

