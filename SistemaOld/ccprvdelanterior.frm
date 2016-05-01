VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCcprvselAnterior 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Proveedores (Selectivo)"
   ClientHeight    =   7365
   ClientLeft      =   450
   ClientTop       =   825
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11100
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia"
      Height          =   540
      Left            =   4560
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   3600
      TabIndex        =   20
      Top             =   4080
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      TabIndex        =   18
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4560
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      _Version        =   327680
      BackColor       =   12640511
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      Height          =   300
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2415
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   3735
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wccprvsel.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   6495
      ItemData        =   "ccprvdelanterior.frx":0000
      Left            =   6840
      List            =   "ccprvdelanterior.frx":0007
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4560
      TabIndex        =   21
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
End
Attribute VB_Name = "PrgCcprvselAnterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private WPorce As Double
Private WTotal As Double
Private WSaldo As Double
Private WSaldoOriginal As Double
Private Vector(1000) As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstCambio As Recordset
Dim spCambio As String
Dim XParam As String
Dim WRetIb As Double
Dim WRetencion As Double
Dim WTipoprv As Integer
Dim WAcumulado As Double
Private XNeto As Double
Private XBruto As Double
Private XIva As Double
Private XTBase As Double
Private WNeto As Double
Private WAnticipo As Double
Private WBruto As Double
Private WIva As Double
Private WRetenido As Double
Private WParametro(0 To 10) As Double
Private WTasa1(10) As Double
Private WTipoiva As Single
Private WTipoIb As Single
Private WTipoIbCaba As Single
Private WFecha As String
Private XImpor As Double
Private WAuxi As Double
Private WAuxi1 As Double
Private WProveedor As String
Private WRete As Double
Dim WVectorIb(1000) As Double
Dim WPorceIb As Double
Dim WControl As String

Dim AcumulaNeto As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String

Private Sub Acepta_Click()

    On Error GoTo WError

    Call Valida_fecha1(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "Formato de Fecha de emision, formato valido : dd/mm/aaaa"
        A% = MsgBox(m$, 0, "Listado de Cuenta Corriente de Proveedores (Selectivo)")
        Exit Sub
    End If
    
    OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    OrdDate = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
    
    If OrdFecha < OrdDate Then
        m$ = "La Fecha de Emision es Menor a la fecha del dia"
        A% = MsgBox(m$, 0, "Listado de Cuenta Corriente de Proveedores (Selectivo)")
        Exit Sub
    End If
    
    spCambio = "ConsultaCambio " + "'" + Fecha.Text + "'"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambio.RecordCount > 0 Then
        ParidadTotal = rstCambio!Cambio
        rstCambio.Close
            Else
        ParidadTotal = 0
        Rem m$ = "La Fecha de Emision no posee paridad informada"
        Rem A% = MsgBox(m$, 0, "Listado de Cuenta Corriente de Proveedores (Selectivo)")
        Rem Exit Sub
    End If
    
    If ParidadTotal = 0 Then
        m$ = "La Fecha de Emision no posee paridad informada"
        A% = MsgBox(m$, 0, "Listado de Cuenta Corriente de Proveedores (Selectivo)")
        Exit Sub
    End If

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores (Selectivo)"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    da = ""
    With rstImpCtaCtePrv
        .Index = "Claveimpre"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    For iRow = 1 To 999
        
        WProveedor = WVector1.TextMatrix(iRow, 1)
        Acumula = 0
        
        If Trim(WProveedor) <> "" Then
        
            XParam = "'" + WProveedor + "','" _
                        + WProveedor + "'"
            spCtaprv = "ListaCtaprvDesdeHasta " + XParam
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
            
                With RstCtaPrv
                
                    .MoveFirst
                
                    If .NoMatch = False Then
        
                    Do
                
                        If WProveedor = !Proveedor And WProveedor <> "" Then
                        
                            WPago = IIf(IsNull(!Pago), "0", !Pago)
                            WParidad = IIf(IsNull(!Paridad), "0", !Paridad)
                            
                            If WPago <> 2 Then
                                WTotal = !Total
                                WSaldo = !Saldo
                                Call Redondeo(WSaldo)
                                WTotalUs = 0
                                WSaldoUs = 0
                                WSaldoOriginal = 0
                                    Else
                                WTotal = (!Total / WParidad) * ParidadTotal
                                WSaldo = (!Saldo / WParidad) * ParidadTotal
                                Call Redondeo(WSaldo)
                                WTotalUs = (!Total / WParidad)
                                WSaldoUs = (!Saldo / WParidad)
                                WSaldoOriginal = !Saldo
                                Call Redondeo(WSaldoOriginal)
                            End If
                    
                            WProveedor = !Proveedor
                            WLetra = !Letra
                            WTipo = !Tipo
                            WPunto = !Punto
                            WNumero = !Numero
                            WFecha = !Fecha
                            WEstado = !Estado
                            Wvencimiento = !Vencimiento
                            WVencimiento1 = !Vencimiento1
                            WNroInterno = !NroInterno
                            WClave = !Clave
                            WOrdFecha = !OrdFecha
                            WOrdVencimiento = !OrdVencimiento
                            WImpre = !Impre
                            WPago = IIf(IsNull(!Pago), "0", !Pago)
                            WParidad = IIf(IsNull(!Paridad), "0", !Paridad)
                            
                            If Tipo2 = True Or WSaldo <> 0 Then
                        
                                With rstImpCtaCtePrv
                                    .Index = "CtaCte"
                                    .AddNew
                                    !Proveedor = WProveedor
                                    !Letra = WLetra
                                    !Tipo = WTipo
                                    !Punto = WPunto
                                    !Numero = WNumero
                                    !Fecha = WFecha
                                    !Estado = WEstado
                                    !Vencimiento = Wvencimiento
                                    !Vencimiento1 = WVencimiento1
                                    !NroInterno = WNroInterno
                                    !Total = WTotal
                                    !Saldo = WSaldo
                                    !SaldoList = WSaldo
                                    !Clave = WClave
                                    !OrdFecha = WOrdFecha
                                    !OrdVencimiento = WOrdVencimiento
                                    !Impre = WImpre
                                    !lista = "S"
                                    Acumula = Acumula + !Saldo
                                    !Acumulado = Acumula
                                    !Titulo = WTitulo
                                    !Pago = WPago
                                    !Paridad = WParidad
                                    !TotalUS = WTotalUs
                                    !SaldoUs = WSaldoUs
                                    !Paridadlistado = ParidadTotal
                                    !SaldoOriginal = WSaldoOriginal
                                    .Update
                                End With
                            
                            End If
                    
                        End If
                    
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                    End If
                End With
                RstCtaPrv.Close
                
            End If
            
        End If
        
    Next iRow
    
    Pasa = 0
    Acumula = 0
    AcumulaII = 0
    AcumulaExento = 0

    With rstImpCtaCtePrv
    
            .Index = "ClaveImpre"
            .MoveFirst
            Do
            
                Rem If !Proveedor > Hasta.Text Then
                Rem    Exit Do
                Rem End If
                
                If Pasa = 0 Then
                    Pasa = 1
                    Acumula = 0
                    AcumulaUs = 0
                    AcumulaII = 0
                    AcumulaNeto = 0
                    AcumulaExento = 0
                    Corte = !Proveedor
                    Erase WVectorIb
                    LugarIb = 0
                End If
                
                If Corte <> !Proveedor Then
                    Acumula = 0
                    AcumulaUs = 0
                    AcumulaII = 0
                    AcumulaNeto = 0
                    AcumulaExento = 0
                    Corte = !Proveedor
                    Erase WVectorIb
                    LugarIb = 0
                End If
                
                .Edit
                !SaldoList = 0
                
                WSaldo = !Saldo
                WSaldoUs = !SaldoUs
                Call Redondeo(WSaldo)
                !SaldoList = WSaldo
                Acumula = Acumula + WSaldo
                !Acumulado = Acumula
                AcumulaUs = AcumulaUs + WSaldoUs
                !AcumulaUs = AcumulaUs
                
                WTotal = !Total
                Call Redondeo(WTotal)
                
                If WTotal = WSaldo Then
                    WPorce = 1
                        Else
                    WPorce = WSaldo / WTotal
                    Call Redondeo(WPorce)
                End If
    
                WProveedor = !Proveedor
                WNroInterno = !NroInterno
                WNombre = ""
                WCheque = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    WCheque = RstProveedor!NombreCheque
                    WTipoIb = RstProveedor!CodIb
                    WTipoIbCaba = RstProveedor!CodIbCaba
                    WTipoiva = RstProveedor!Iva
                    WTipoprv = Val(RstProveedor!Tipo) + 1
                    WPorceIb = IIf(IsNull(RstProveedor!PorceIb), "0", RstProveedor!PorceIb)
                    WPorceIbCaba = IIf(IsNull(RstProveedor!PorceIbCaba), "0", RstProveedor!PorceIbCaba)
                    RstProveedor.Close
                End If
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                ZRechazado = 0
                ZLetra = ""
                ZNeto = 0
                ZIva = 0
                ZIva5 = 0
                ZIva27 = 0
                ZIva105 = 0
                ZIb = 0
                ZExento = 0
                ZTotal = 0
                
                ZFactura = ""
                
                spIvaComp = "Consultaivacomp " + "'" + Str$(WNroInterno) + "'"
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
                    ZLetra = rstIvaComp!Letra
                    ZNeto = rstIvaComp!Neto
                    ZIva = rstIvaComp!Iva21
                    ZIva5 = rstIvaComp!Iva5
                    ZIva27 = rstIvaComp!Iva27
                    ZIva105 = IIf(IsNull(rstIvaComp!Iva105), "0", rstIvaComp!Iva105)
                    ZIb = rstIvaComp!Ib
                    ZExento = rstIvaComp!Exento
                    ZTotal = ZNeto + ZIva + ZIva27 + ZIva105 + ZIb + ZIva5 + ZExento
                    ZFactura = "S"
                    rstIvaComp.Close
                End If
                
                
                
                WSaldoOriginal = !SaldoOriginal
                
                If ZRechazado = 1 Then
                    AcumulaExento = AcumulaExento + WSaldo
                End If
                
                WRetIb = 0
                WRetIva = 0
                WRetgan = 0
                WAcumulaIb = 0
                
                ZZEntraVec = "S"
    
                If WTipoIb = 0 Or WTipoIb = 1 Then
                
                    If ZRechazado = 0 Then
                        If WSaldoOriginal > 0 Then
                            LugarIb = LugarIb + 1
                            WVectorIb(LugarIb) = WSaldoOriginal
                            LugarIb = LugarIb + 1
                            WVectorIb(LugarIb) = WSaldo - WSaldoOriginal
                                Else
                            LugarIb = LugarIb + 1
                            WVectorIb(LugarIb) = WSaldo
                        End If
                        ZZEntraVec = "N"
                    End If
                    
                    XBruto = !Acumulado
                    If WTipoiva = 2 Then
                        XNeto = (XBruto / 1.21)
                            Else
                        XNeto = XBruto
                    End If
                    XIva = XBruto - XNeto
                    XTBase = XNeto
            
                    For CicloIb = 1 To LugarIb
                    
                        XBruto = WVectorIb(CicloIb)
                        If WTipoiva = 2 Then
                            XNeto = (XBruto / 1.21)
                                Else
                            XNeto = XBruto
                        End If
                        XIva = XBruto - XNeto
                        XTBase = XNeto
                        Call Redondeo(XTBase)
                    
                        Select Case WTipoIb
                            Case 0, 1
                                WRete = XTBase * (WPorceIb / 100)
                                Call Redondeo(WRete)
                                WAcumulaIb = WAcumulaIb + WRete
                            Case Else
                        End Select

                        Rem Select Case WTipoIb
                        Rem     Case 0
                        Rem         WRete = XTBase * (0.75 / 100)
                        Rem         Call Redondeo(WRete)
                        Rem         WAcumulaIb = WAcumulaIb + WRete
                        Rem     Case 1
                        Rem         WRete = XTBase * (1.75 / 100)
                        Rem         Call Redondeo(WRete)
                        Rem         WAcumulaIb = WAcumulaIb + WRete
                        Rem     Case Else
                        Rem         WAcumulaIb = 0
                        Rem End Select
                        
                    Next CicloIb
                    
                    WRetIb = WAcumulaIb
                    Call Redondeo(WRetIb)
        
                End If
                
                If WTipoIbCaba = 3 Or WTipoIbCaba = 4 Or WPorceIbCaba <> 0 Then
                If WTipoIbCaba <> 2 Then
                    
                    If Val(Wempresa) = 1 Then
            
                        If ZZEntraVec = "S" Then
                            If ZRechazado = 0 Then
                                If WSaldoOriginal > 0 Then
                                    LugarIb = LugarIb + 1
                                    WVectorIb(LugarIb) = WSaldoOriginal
                                    LugarIb = LugarIb + 1
                                    WVectorIb(LugarIb) = WSaldo - WSaldoOriginal
                                        Else
                                    LugarIb = LugarIb + 1
                                    WVectorIb(LugarIb) = WSaldo
                                End If
                            End If
                        End If
                        
                        XBruto = !Acumulado
                        If WTipoiva = 2 Then
                            XNeto = (XBruto / 1.21)
                                Else
                            XNeto = XBruto
                        End If
                        XIva = XBruto - XNeto
                        XTBase = XNeto
                
                        Rem If XTBase >= 500 Then
                        
                            For CicloIb = 1 To LugarIb
                            
                                XBruto = WVectorIb(CicloIb)
                                If WTipoiva = 2 Then
                                    XNeto = (XBruto / 1.21)
                                        Else
                                    XNeto = XBruto
                                End If
                                XIva = XBruto - XNeto
                                XTBase = XNeto
                                Call Redondeo(XTBase)
                                If XTBase < 300 Then
                                    XTBase = 0
                                End If
                                
                                If WPorceIbCaba <> 0 Then
                                    WRete = XTBase * (WPorceIbCaba / 100)
                                    Call Redondeo(WRete)
                                        Else
                                    If WTipoIbCaba = 3 Then
                                        WRete = XTBase * (3 / 100)
                                            Else
                                        WRete = XTBase * (4.5 / 100)
                                    End If
                                    Call Redondeo(WRete)
                                End If
                                WAcumulaIb = WAcumulaIb + WRete
                                
                            Next CicloIb
                            
                            WRetIb = WAcumulaIb
                            
                        Rem End If
                                
                        Call Redondeo(WRetIb)
                    
                    End If
                    End If
                    
                End If
            
                WAcumulado = !Acumulado - AcumulaExento
                Call calculaRetencion
                WRetgan = WRetencion
                            
                                
                Rem suma a la retencion de ib la retenciones
                Rem de iva de las facturas M
                If ZLetra = "M" Then
                    If ZNeto >= 1000 Then
                        WRetIb = WRetIb + ZIva
                    End If
                End If
                            
                !Acuneto = !Acumulado - WRetIb - WRetgan
                !Nombre = WNombre
                !Cheque = WCheque
                !ReteIb = WRetIb
                !ReteGan = WRetgan
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With

   If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
    Listado.Action = 1
    
    WVector1.Col = 1
    WVector1.Row = 1
    Exit Sub
        
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    With rstEmpresa
        .Close
    End With
    With rstImpCtaCtePrv
        .Close
    End With
    PrgCcprvsel.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImpCtaCtePrv
End Sub



Private Sub Form_Load()

    Call Limpia_Vector
 
    Panta.Value = False
    Impresora.Value = True
    Tipo1.Value = True
    Tipo2.Value = False
    Frame2.Visible = True

    WParametro(0) = 0
    WParametro(1) = 2000
    WParametro(2) = 4000
    WParametro(3) = 8000
    WParametro(4) = 14000
    WParametro(5) = 24000
    WParametro(6) = 1000000
    
    WTasa1(1) = 0.1
    WTasa1(2) = 0.14
    WTasa1(3) = 0.18
    WTasa1(4) = 0.22
    WTasa1(5) = 0.26
    WTasa1(6) = 0.26
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Limpia_Click()
    Call Limpia_Vector
End Sub

Private Sub Proceso_Click()

    Rem With rstProceso1
    Rem     .Index = "Numero"
    Rem     .Seek ">=", da
    Rem     If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             WProveedor = !Proveedor
    Rem
    Rem             With rstProveedor
    Rem                 .Index = "Proveedor"
    Rem                 .Seek "=", WProveedor
    Rem                 If .NoMatch = False Then
    Rem                     WNombre = !Nombre
    Rem                 End If
    Rem             End With
    Rem
    Rem
    Rem             Lugar1 = Int(!Numero / 10)
    Rem             Lugar2 = !Numero - Lugar1
    Rem
    Rem             DBGrid1.FirstRow = Lugar1
    Rem             DBGrid1.Row = Lugar2 - 1
    Rem             DBGrid1.Col = 0
    Rem             DBGrid1.Text = WProveedor
    Rem             DBGrid1.Col = 1
    Rem             DBGrid1.Text = WNombre
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End If
    Rem End With

    WVector1.Col = 1
    WVector1.Row = 1


End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    Indice = Pantalla.ListIndex
    
    Ingre = "S"
    For A = 1 To 1000
        If WVector1.TextMatrix(A, 1) = WIndice.List(Indice) Then
            Ingre = "N"
        End If
    Next A
    
    If Ingre = "S" Then
    
        For Ciclo = 1 To 1000
            If Trim(WVector1.TextMatrix(Ciclo, 1)) = "" Then
                WVector1.Row = Ciclo
                WVector1.Col = 1
                WVector1.Text = WIndice.List(Indice)
                WTexto1.Text = WIndice.List(Indice)
                Call WTexto1_KeyDown(13, 0)
                Exit For
            End If
        Next Ciclo
        
    End If
    
End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If Ayuda.Text <> "" Then
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Else
        spProveedor = "ListaProveedoresOrdConsulta"
    End If
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                    
                    
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    RstProveedor.Close
    
    End If
    
    End If

End Sub

Private Sub calculaRetencion()

    WRetencion = 0
    
    If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Or WTipoprv = 6 Or WTipoprv = 7 Then
        
            XBruto = WAcumulado
            If WTipoiva = 2 Then
                XNeto = (XBruto / 1.21)
                    Else
                XNeto = XBruto
            End If
            XIva = XBruto - XNeto
            XTBase = XNeto
            
            WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
            
            ClaveRetencion = WFecha + WProveedor
            spRetencion = "ConsultaRetencion " + "'" + ClaveRetencion + "'"
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            If rstRetencion.RecordCount > 0 Then
                WNeto = rstRetencion!Neto
                WAnticipo = rstRetencion!Anticipo
                WBruto = rstRetencion!Bruto
                WIva = rstRetencion!Iva
                WRetenido = rstRetencion!Retenido
                rstRetencion.Close
                    Else
                XFecha = WFecha
                XProveedor = WProveedor
                XXNeto = ""
                XXAnticipo = ""
                XXBruto = ""
                XXIva = ""
                XXRetenido = ""
                XClave = XFecha + XProveedor
                
                XParam = "'" + XClave + "','" _
                        + XFecha + "','" + XProveedor + "','" _
                        + XXNeto + "','" _
                        + XXRetenido + "','" + XXAnticipo + "','" _
                        + XXBruto + "','" _
                        + XXAcumulado + "'"
                    
                spRstRetencion = "AltaRetencion " + XParam
                Set RstRstRetencion = db.OpenRecordset(spRstRetencion, dbOpenSnapshot, dbSQLPassThrough)
                    
                WNeto = 0
                WAnticipo = 0
                WBruto = 0
                WIva = 0
                WRetenido = 0
            End If

            Select Case WTipoprv
                Case 1
                    WMinimo = 12000
                Case 2
                    WMinimo = 1200
                Case 3
                    WMinimo = 1200
                Case 6
                    WMinimo = 5000
                Case 7
                    WMinimo = 6500
                Case Else
            End Select

            WAcupag = WNeto + XTBase
            WAuxi = WAcupag - WMinimo

            If WAuxi <= 0 Then
                WAuxi = 0
                WRetencion = 0
            End If

            WTasa = 0.02
            If WTipoprv = 1 Then
                    WTasa = 0.02
            End If
            If WTipoprv = 3 Then
                    WTasa = 0.06
            End If
            If WTipoprv = 7 Then
                    WTasa = 0.0025
            End If

            Select Case WTipoprv
                Case 2
                    WRetencion = 0
                    WTope = 0
                    WTope1 = 0
                    
                    For da = 0 To 5
                        If WAuxi >= WParametro(da) And WAuxi < WParametro(da + 1) Then
                            WTope1 = WAuxi
                            WTope = WParametro(da)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(da + 1)
                            WRetencion = WRetencion + WSum
                        End If
                        If WAuxi >= WParametro(da + 1) Then
                            WTope1 = WParametro(da + 1)
                            WTope = WParametro(da)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(da + 1)
                            WRetencion = WRetencion + WSum
                        End If
                    Next da
                    
                Case Else
                    WRetencion = WAuxi * WTasa
                    
            End Select

            WRetencion = WRetencion - WRetenido

            If WRetencion < 20 Then
                WRetencion = 0
                        Else
                If WRetencion > XNeto Then
                        WRetencion = 0
                End If
            End If
                    
            Call Redondeo(WRetencion)

    End If

End Sub


Rem
Rem Controles de la grilla
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub


Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 1, 2
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            For A = 1 To 1000
                If WVector1.TextMatrix(A, 1) = WVector1.Text And A <> WVector1.Row Then
                    WControl = "N"
                End If
            Next A
            
            If WControl = "S" Then
                spProveedor = "ConsultaProveedores " + "'" + WVector1.Text + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = RstProveedor!Nombre
                    WVector1.Col = 1
                    RstProveedor.Close
                        Else
                    WControl = "N"
                End If
            End If
            
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Proveedor"
                WVector1.ColWidth(Ciclo) = 1600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 11
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi2 = WVector1.Text
        If WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub


