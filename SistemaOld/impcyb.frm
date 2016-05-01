VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImpcyb 
   Caption         =   "Listado de Imputaciones de Caja Y Banco"
   ClientHeight    =   7365
   ClientLeft      =   2970
   ClientTop       =   525
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   5655
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4680
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1815
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4455
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox DesdeCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   19
         Text            =   " "
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox HastaCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   18
         Text            =   " "
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox TipoList 
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Text            =   " "
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   360
         TabIndex        =   13
         Top             =   2640
         Width           =   3135
         Begin VB.CheckBox Tipo3 
            Caption         =   "Recibos"
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Tipo2 
            Caption         =   "Depositos"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Tipo1 
            Caption         =   "Pagos"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   282
         Left            =   1320
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
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wimpcyb.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgImpcyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XBanco(100) As String
Private XProveedor(10000) As String
Dim EntraRec(100000, 21) As String
Dim LugarRec As Single
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstDepositos As Recordset
Dim spDepositos As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo Error_Programa
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
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
    
    XLugar = 0
    Erase XBanco
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    XLugar = XLugar + 1
                    XBanco(rstBanco!Banco) = rstBanco!Cuenta
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
    
    XLugarProve = 0
    Erase XProveedor
    spProveedor = "ListaProveedores"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    If Val(RstProveedor!Provincia) = 24 Then
                        XLugarProve = XLugarProve + 1
                        XProveedor(XLugarProve) = RstProveedor!Proveedor
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
    

    If Tipo1.Value = 1 Then
    
        XParam = "'" + WDesde + "','" _
                     + WHasta + " '"
    
        spPagos = "ListaPagosFecha " + XParam
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
    
            With rstPagos
                    .MoveFirst
                    Do
                        If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                        
                            If Corte <> !Orden Then
                                Corte = !Orden
                                Renglon = 0
                            End If
0
                            Select Case Val(!Tiporeg)
                                Case 1
                                    If !TipoOrd = "3" Or !TipoOrd = "4" Or !TipoOrd = "5" Then
                                    
                                        Select Case !TipoOrd
                                            Case "4"
                                                WCuenta = XBanco(!Banco2)
                                            Case "5"
                                                WCuenta = "111"
                                            Case Else
                                                WCuenta = !Cuenta
                                        End Select
                                        
                                            Else
                                            
                                        Rem proveedor
                                        WCuenta = "2001"
                                        If Val(WEmpresa) = 8 And "10077777777" = !Proveedor Then
                                            WCuenta = "2046"
                                        End If
                                        For WDa = 1 To XLugarProve
                                            If XProveedor(WDa) = !Proveedor Then
                                                WCuenta = "2010"
                                                Exit For
                                            End If
                                        Next WDa
                                        
                                        
                                    End If
                                    
                                    ZZZValida = "S"
                                    ZZZLetra = !Letra1
                                    ZZZTipo = !Tipo1
                                    ZZZPunto = !Punto1
                                    ZZZNumero = !Numero1
                                    If Trim(ZZZLetra) = "" And Val(ZZZTipo) = 0 And Val(ZZZPunto) = 0 And Val(ZZZNumero) = 0 Then
                                        If Trim(!Cuenta) <> "" And !Cuenta <> "999999" Then
                                            WCuenta = !Cuenta
                                        End If
                                    End If
                                    
                                    WImporte = !Importe1
                                    WProveedor = !Proveedor
                                    WOrden = !Orden
                                    WTipo = !Tipo1
                                    WLetra = !Letra1
                                    WPunto = !Punto1
                                    WNumero = !Numero1
                                    WFecha = !Fecha
                                    WObservaciones = !Observaciones
                                    
                                    With rstImpcyb
                                        .Index = "Clave"
                                        .AddNew
                                        !Tipomovi = "1"
                                        !NroInterno = WOrden
                                        !Proveedor = WProveedor
                                        !TipoComp = WTipo
                                        !LetraComp = WLetra
                                        !PuntoComp = WPunto
                                        !NroComp = WNumero
                                        Renglon = Renglon + 1
                                        Auxi1 = Str$(Renglon)
                                        Call Ceros(Auxi1, 2)
                                        !Renglon = 0
                                        !Fecha = WFecha
                                        !Observaciones = WObservaciones
                                        !Cuenta = WCuenta
                                        !Debito = WImporte
                                        !Credito = 0
                                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        !Titulo = "Pagos"
                                        !Empresa = 1
                                        !Clave = !Tipomovi + !NroInterno + !Renglon
                                        !Titulolist = WTitulo
                                        !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                        !ClaveOrd = !Tipomovi + !NroInterno
                                        .Update
                                
                                    End With
                                    
                                    If Val(!Renglon) = 1 And !RetOtra <> 0 Then
                                    
                                        WImporte = !RetOtra
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2108"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                    If Val(!Renglon) = 1 And !Retencion <> 0 Then
                                    
                                        WImporte = !Retencion
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2101"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                    If Val(!Renglon) = 1 And !RetIva <> 0 Then
                                    
                                        WImporte = !RetIva
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2111"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                    ZZRetIbCiudad = IIf(IsNull(!RetIbCiudad), "", !RetIbCiudad)
                                    If Val(!Renglon) = 1 And ZZRetIbCiudad <> 0 Then
                                    
                                        WImporte = ZZRetIbCiudad
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2113"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                                                    
                                Case Else
                                    Select Case Val(!Tipo2)
                                        Case 1
                                            Rem caja
                                            WCuenta = "1"
                                        Case 2
                                            Rem banco
                                            WCuenta = "999999"
                                            WBanco2 = !Banco2
                                            WCuenta = XBanco(WBanco2)
                                        Case 3
                                            Rem che ter
                                            WCuenta = "40"
                                        Case 5
                                            Rem U$S
                                            WCuenta = "2"
                                        Case 6
                                            Rem Varios
                                            WCuenta = !Cuenta
                                        Case 7
                                            Rem Patacones
                                            WCuenta = "7"
                                        Case 8
                                            Rem Lecop
                                            WCuenta = "8"
                                        Case Else
                                            Rem documentos
                                            WCuenta = "101"
                                    End Select
                                            
                                    WImporte = !Importe2
                                    WProveedor = !Proveedor
                                    WOrden = !Orden
                                    WTipo = !Tipo2
                                    WNumero = !Numero2
                                    WFecha = !Fecha
                                    WObservaciones = !Observaciones
                                    
                                    With rstImpcyb
                                        .Index = "Clave"
                                        .AddNew
                                        !Tipomovi = "1"
                                        !NroInterno = WOrden
                                        !Proveedor = WProveedor
                                        !TipoComp = WTipo
                                        !LetraComp = ""
                                        !PuntoComp = ""
                                        !NroComp = WNumero
                                        Renglon = Renglon + 1
                                        Auxi1 = Str$(Renglon)
                                        Call Ceros(Auxi1, 2)
                                        !Renglon = Auxi1$
                                        !Fecha = WFecha
                                        !Observaciones = WObservaciones
                                        !Cuenta = WCuenta
                                        !Credito = WImporte
                                        !Debito = 0
                                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        !Titulo = "Pagos"
                                        !Empresa = 1
                                        !Clave = !Tipomovi + !NroInterno + !Renglon
                                        !Titulolist = WTitulo
                                        !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                        !ClaveOrd = !Tipomovi + !NroInterno
                                        .Update
                                
                                    End With
                                    
                                    If Val(!Renglon) = 1 And !Retencion <> 0 Then
                                    
                                        WImporte = !Retencion
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2101"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                    If Val(!Renglon) = 1 And !RetOtra <> 0 Then
                                    
                                        WImporte = !RetOtra
                                        WProveedor = !Proveedor
                                        WOrden = !Orden
                                        WFecha = !Fecha
                                        WObservaciones = !Observaciones
                                    
                                        With rstImpcyb
                                            .Index = "Clave"
                                            .AddNew
                                            !Tipomovi = "1"
                                            !NroInterno = WOrden
                                            !Proveedor = WProveedor
                                            !TipoComp = ""
                                            !LetraComp = ""
                                            !PuntoComp = ""
                                            !NroComp = ""
                                            Renglon = Renglon + 1
                                            Auxi1 = Str$(Renglon)
                                            Call Ceros(Auxi1, 2)
                                            !Renglon = Auxi1$
                                            !Fecha = WFecha
                                            !Observaciones = WObservaciones
                                            !Cuenta = "2108"
                                            !Credito = WImporte
                                            !Debito = 0
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Titulo = "Pagos"
                                            !Empresa = 1
                                            !Clave = !Tipomovi + !NroInterno + !Renglon
                                            !Titulolist = WTitulo
                                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                            !ClaveOrd = !Tipomovi + !NroInterno
                                            .Update
                                    
                                        End With
                                    End If
                                    
                                    
                                End Select
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
            End With
            rstPagos.Close
            
        End If
    
    End If
    
    If Tipo2.Value = 1 Then
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + " '"
                 
    spDepositos = "ListaDepositosFecha " + XParam
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
    
    With rstDepositos
            .MoveFirst
            Do
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                
                    WDeposito = !Deposito
                    WTipo = !Tipo2
                    WNumero = !Numero2
                    WImporte = !Importe2
                    WFecha = !Fecha
                    WBanco = !Banco
                            
                    With rstImpcyb
                    
                        .Index = "Clave"
                        
                        .AddNew
                        
                        WCuenta = XBanco(WBanco)
                        !Tipomovi = "2"
                        !NroInterno = WDeposito
                        !Proveedor = ""
                        !TipoComp = WTipo
                        !LetraComp = ""
                        !PuntoComp = ""
                        !NroComp = WNumero
                        Rem Renglon = Renglon + 1
                        Rem Auxi1 = Str$(Renglon)
                        Rem Call Ceros(Auxi1, 2)
                        !Renglon = 0
                        !Fecha = WFecha
                        !Observaciones = ""
                        !Cuenta = WCuenta
                        !Debito = WImporte
                        !Credito = 0
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Titulo = "Depositos"
                        !Empresa = 1
                        !Clave = !Tipomovi + !NroInterno + !Renglon
                        !Titulolist = WTitulo
                        !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                        !ClaveOrd = !Tipomovi + !NroInterno
                        .Update
                        
                        .AddNew
                        
                        Select Case Val(WTipo)
                            Case 2
                                WCuenta = "2"
                            Case 3
                                WCuenta = "40"
                            Case Else
                                WCuenta = "1"
                        End Select
                        
                        !Tipomovi = "2"
                        !NroInterno = WDeposito
                        !Proveedor = ""
                        !TipoComp = WTipo
                        !LetraComp = ""
                        !PuntoComp = ""
                        !NroComp = WNumero
                        Rem Renglon = Renglon + 1
                        Rem Auxi1 = Str$(Renglon)
                        Rem Call Ceros(Auxi1, 2)
                        !Renglon = 0
                        !Fecha = WFecha
                        !Observaciones = ""
                        !Cuenta = WCuenta
                        !Debito = 0
                        !Credito = WImporte
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Titulo = "Depositos"
                        !Empresa = 1
                        !Clave = !Tipomovi + !NroInterno + !Renglon
                        !Titulolist = WTitulo
                        !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                        !ClaveOrd = !Tipomovi + !NroInterno
                        .Update
                        
                    End With
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstDepositos.Close
    
    End If
    
    End If
    
    If Tipo3.Value = 1 Then


    Erase EntraRec
    LugarRec = 0



    XParam = "'" + WDesde + "','" _
                 + WHasta + " '"
                 
    spRecibos = "ListaRecibosFecha " + XParam
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

    With rstRecibos
            .MoveFirst
            Do
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                
                    LugarRec = LugarRec + 1
                    
                    EntraRec(LugarRec, 1) = !FechaOrd
                    EntraRec(LugarRec, 2) = !Recibo
                    EntraRec(LugarRec, 3) = !Tiporeg
                    EntraRec(LugarRec, 4) = !TipoRec
                    EntraRec(LugarRec, 5) = !Cuenta
                    EntraRec(LugarRec, 6) = !Tipo1
                    EntraRec(LugarRec, 7) = !Cliente
                    EntraRec(LugarRec, 8) = Str$(!Importe1)
                    EntraRec(LugarRec, 9) = Str$(!Paridad)
                    EntraRec(LugarRec, 10) = !Letra1
                    EntraRec(LugarRec, 11) = !Punto1
                    EntraRec(LugarRec, 12) = !Numero1
                    EntraRec(LugarRec, 13) = !Fecha
                    EntraRec(LugarRec, 14) = !Tipo2
                    EntraRec(LugarRec, 15) = Str$(!Importe2)
                    EntraRec(LugarRec, 16) = !Numero2
                    EntraRec(LugarRec, 17) = Str$(!Retganancias)
                    EntraRec(LugarRec, 18) = !Renglon
                    EntraRec(LugarRec, 19) = Str$(!RetIva)
                    EntraRec(LugarRec, 20) = Str$(!RetOtra)
                    EntraRec(LugarRec, 21) = Str$(!RetSuss)
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstRecibos.Close
    
    End If
    
    
    
    Corte = ""
        
    For CiclaRec = 1 To LugarRec
    
                    ZFechaOrd = EntraRec(CiclaRec, 1)
                    ZRecibo = EntraRec(CiclaRec, 2)
                    
                    
                    Ztiporeg = EntraRec(CiclaRec, 3)
                    Ztiporec = EntraRec(CiclaRec, 4)
                    ZCuenta = EntraRec(CiclaRec, 5)
                    ZTipo1 = EntraRec(CiclaRec, 6)
                    ZCliente = EntraRec(CiclaRec, 7)
                    ZImporte1 = Val(EntraRec(CiclaRec, 8))
                    ZParidad = Val(EntraRec(CiclaRec, 9))
                    ZLetra1 = EntraRec(CiclaRec, 10)
                    ZPunto1 = EntraRec(CiclaRec, 11)
                    ZNumero1 = EntraRec(CiclaRec, 12)
                    ZFecha = EntraRec(CiclaRec, 13)
                    ZTipo2 = EntraRec(CiclaRec, 14)
                    ZImporte2 = Val(EntraRec(CiclaRec, 15))
                    ZNumero2 = EntraRec(CiclaRec, 16)
                    ZRetganancias = Val(EntraRec(CiclaRec, 17))
                    ZRenglon = EntraRec(CiclaRec, 18)
                    ZRetIva = Val(EntraRec(CiclaRec, 19))
                    ZRetOtra = Val(EntraRec(CiclaRec, 20))
                    ZRetSuss = Val(EntraRec(CiclaRec, 21))
                
                    If Corte <> ZRecibo Then
                        Corte = ZRecibo
                        Renglon = 0
                    End If
                
                    Select Case Val(Ztiporeg)
                        Case 1
                            
                            WProv = ""
                            WCliente = ZCliente
                            spClientes = "ConsultaClientes " + "'" + WCliente + "'"
                            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                            If rstClientes.RecordCount > 0 Then
                                WProv = rstClientes!Provincia
                                rstClientes.Close
                            End If
                            
                            If Ztiporec = "3" Then
                                WCuenta = ZCuenta
                                    Else
                                Rem clientes
                                If Val(ZTipo1) > 49 Then
                                    WCuenta = "101"
                                        Else
                                    WCuenta = "91"
                                    If Val(WEmpresa) <> 1 Then
                                        If Val(WProv) = 24 Then
                                            WCuenta = "92"
                                        End If
                                    End If
                                End If
                            End If
                            
                            If Val(WProv) = 24 Then
                            
                                Auxi1 = ZRecibo
                                Call Ceros(Auxi1, 8)
    
                                ClaveCtacte = "06" + Auxi1 + "01"
                                spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                                If rstCtaCte.RecordCount > 0 Then
                                    ZParidad = rstCtaCte!Paridad
                                    rstCtaCte.Close
                                        Else
                                    ClaveCtacte = "07" + Auxi1 + "01"
                                    spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstCtaCte.RecordCount > 0 Then
                                        ZParidad = (rstCtaCte!Paridad)
                                        rstCtaCte.Close
                                    End If
                                End If
                            
                                If Val(ZTipo1) <> 7 Then
                                    WImporte = ZImporte1 / ZParidad
                                        Else
                                    WImporte = ZImporte1
                                End If
                                
                                If Val(ZTipo1) <> 7 Then
                                    With rstCtaCte
                                        ClaveCtacte = ZTipo1 + ZNumero1 + "01"
                                        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstCtaCte.RecordCount > 0 Then
                                            If rstCtaCte!TotalUS <> 0 Then
                                                Pari = rstCtaCte!Paridad
                                                WImporte = WImporte * Pari
                                            End If
                                            rstCtaCte.Close
                                        End If
                                    End With
                                End If
                                
                                    Else
                                WImporte = ZImporte1
                            End If
                            
                            WCliente = ZCliente
                            WRecibo = ZRecibo
                            WTipo = ZTipo1
                            WLetra = ZLetra1
                            WPunto = ZPunto1
                            WNumero = ZNumero1
                            WFecha = ZFecha
                            XCuenta = ZCuenta
                            
                            With rstImpcyb
                                .Index = "Clave"
                                .AddNew
                                !Tipomovi = "3"
                                !NroInterno = WRecibo
                                !Proveedor = ""
                                !TipoComp = WTipo
                                !LetraComp = WLetra
                                !PuntoComp = WPunto
                                !NroComp = WNumero
                                Renglon = Renglon + 1
                                Auxi1 = Str$(Renglon)
                                Call Ceros(Auxi1, 2)
                                !Renglon = 0
                                !Fecha = WFecha
                                Rem !Observaciones = WObservaciones
                                !Observaciones = ""
                                !Cuenta = WCuenta
                                !Debito = 0
                                !Credito = WImporte
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Titulo = "Recibos"
                                !Empresa = 1
                                !Clave = !Tipomovi + !NroInterno + !Renglon
                                !Titulolist = WTitulo
                                !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                !ClaveOrd = !Tipomovi + !NroInterno
                                .Update
                        
                            End With
                                                            
                        Case Else
                            Select Case Val(ZTipo2)
                                Case 1
                                    Rem caja
                                    WCuenta = "1"
                                Case 2
                                    Rem cheques
                                    WCuenta = "40"
                                Case 4
                                    WCuenta = ZCuenta
                                Case Else
                                    Rem documentos
                                    WCuenta = "101"
                            End Select
                                    
                            WImporte = ZImporte2
                            WProveedor = ""
                            WRecibo = ZRecibo
                            WTipo = ZTipo2
                            WNumero = ZNumero2
                            WFecha = ZFecha
                            XCuenta = ZCuenta
                            
                            With rstImpcyb
                                .Index = "Clave"
                                .AddNew
                                !Tipomovi = "3"
                                !NroInterno = WRecibo
                                !Proveedor = ""
                                !TipoComp = WTipo
                                !LetraComp = ""
                                !PuntoComp = ""
                                !NroComp = WNumero
                                Renglon = Renglon + 1
                                Auxi1 = Str$(Renglon)
                                Call Ceros(Auxi1, 2)
                                !Renglon = Auxi1$
                                !Fecha = WFecha
                                !Observaciones = ""
                                !Cuenta = WCuenta
                                !Debito = WImporte
                                !Credito = 0
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Titulo = "Recibos"
                                !Empresa = 1
                                !Clave = !Tipomovi + !NroInterno + !Renglon
                                !Titulolist = WTitulo
                                !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                !ClaveOrd = !Tipomovi + !NroInterno
                                .Update
                        
                            End With
                    End Select
                
                    If Val(ZRenglon) = 1 And Val(ZRetganancias) <> 0 Then
                        WImporte = ZRetganancias
                        WProveedor = ""
                        WRecibo = ZRecibo
                        WFecha = ZFecha
                            
                        With rstImpcyb
                            .Index = "Clave"
                            .AddNew
                            !Tipomovi = "3"
                            !NroInterno = WRecibo
                            !Proveedor = ""
                            !TipoComp = ""
                            !LetraComp = ""
                            !PuntoComp = ""
                            !NroComp = ""
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            !Renglon = Auxi1$
                            !Fecha = WFecha
                            !Observaciones = ""
                            !Cuenta = "142"
                            !Debito = WImporte
                            !Credito = 0
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Titulo = "Recibos"
                            !Empresa = 1
                            !Clave = !Tipomovi + !NroInterno + !Renglon
                            !Titulolist = WTitulo
                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                            !ClaveOrd = !Tipomovi + !NroInterno
                            .Update
                        
                        End With
                    End If
                
                    If Val(ZRenglon) = 1 And Val(ZRetIva) <> 0 Then
                        WImporte = ZRetIva
                        WProveedor = ""
                        WRecibo = ZRecibo
                        WFecha = ZFecha
                            
                        With rstImpcyb
                            .Index = "Clave"
                            .AddNew
                            !Tipomovi = "3"
                            !NroInterno = WRecibo
                            !Proveedor = ""
                            !TipoComp = ""
                            !LetraComp = ""
                            !PuntoComp = ""
                            !NroComp = ""
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            !Renglon = Auxi1$
                            !Fecha = WFecha
                            !Observaciones = ""
                            !Cuenta = "153"
                            !Debito = WImporte
                            !Credito = 0
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Titulo = "Recibos"
                            !Empresa = 1
                            !Clave = !Tipomovi + !NroInterno + !Renglon
                            !Titulolist = WTitulo
                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                            !ClaveOrd = !Tipomovi + !NroInterno
                            .Update
                        
                        End With
                    End If
                
                    If Val(ZRenglon) = 1 And Val(ZRetOtra) <> 0 Then
                        WImporte = ZRetOtra
                        WProveedor = ""
                        WRecibo = ZRecibo
                        WFecha = ZFecha
                                
                        With rstImpcyb
                            .Index = "Clave"
                            .AddNew
                            !Tipomovi = "3"
                            !NroInterno = WRecibo
                            !Proveedor = ""
                            !TipoComp = ""
                            !LetraComp = ""
                            !PuntoComp = ""
                            !NroComp = ""
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            !Renglon = Auxi1$
                            !Fecha = WFecha
                            !Observaciones = ""
                            !Cuenta = "161"
                            !Debito = WImporte
                            !Credito = 0
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Titulo = "Recibos"
                            !Empresa = 1
                            !Clave = !Tipomovi + !NroInterno + !Renglon
                            !Titulolist = WTitulo
                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                            !ClaveOrd = !Tipomovi + !NroInterno
                            .Update
                        
                        End With
                    End If
                    
                    If Val(ZRenglon) = 1 And Val(ZRetSuss) <> 0 Then
                        WImporte = ZRetSuss
                        WProveedor = ""
                        WRecibo = ZRecibo
                        WFecha = ZFecha
                                
                        With rstImpcyb
                            .Index = "Clave"
                            .AddNew
                            !Tipomovi = "3"
                            !NroInterno = WRecibo
                            !Proveedor = ""
                            !TipoComp = ""
                            !LetraComp = ""
                            !PuntoComp = ""
                            !NroComp = ""
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            !Renglon = Auxi1$
                            !Fecha = WFecha
                            !Observaciones = ""
                            !Cuenta = "145"
                            !Debito = WImporte
                            !Credito = 0
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Titulo = "Recibos"
                            !Empresa = 1
                            !Clave = !Tipomovi + !NroInterno + !Renglon
                            !Titulolist = WTitulo
                            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                            !ClaveOrd = !Tipomovi + !NroInterno
                            .Update
                        
                        End With
                    End If
                
    Next CiclaRec
    
    
    
    
    
    
    
    End If
    
    
    
    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
                If Val(!Cuenta) < Val(DesdeCuenta.Text) Or Val(!Cuenta) > Val(HastaCuenta.Text) Then
                    .Delete
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
                .Edit
                    
                WCuenta = !Cuenta
                WNombre = ""
                
                spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WNombre = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
                
                !Nombre = WNombre
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With

    Rem Listado.GroupSelectionFormula = "{Impcyb.banco} in " + Chr$(34) + DesdeBanco + Chr$(34) + " to " + Chr$(34) + HastaBanco + Chr$(34) + " and {Pagos.Renglon} = " + Chr$(34) + "01" + Chr$(34)
    Rem Listado.GroupSelectionFormula = "{Impcyb.banco} in 0 to 9999"

    If TipoList.ListIndex = 0 Then
        listado.ReportFileName = "wimpcyb.rpt"
        Rem Listado.ReportFileName = "Salcyb.rpt"
            Else
        listado.ReportFileName = "wimpcyb1.rpt"
    End If

    If Impresora.Value = True Then
        listado.Destination = 1
            Else
        listado.Destination = 0
    End If
    
    listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    listado.Action = 1
    
    Exit Sub
    
    
Error_Programa:
     Rem coderr = Err
     Rem Call Errores(coderr, "Error en el sistema", "Se produjo el error " + Str$(coderr))
     Resume Next
    
End Sub

Private Sub Cancela_Click()
    With rstImpcyb
        .Close
    End With
    Desde.SetFocus
    PrgImpcyb.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Impcyb
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeCuenta.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCuenta.SetFocus
    End If
End Sub

Private Sub HastaCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    spCuenta = "ListaCuentas"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCuenta
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstCuenta!Cuenta
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstCuenta.Close
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WCuenta = WIndice.List(Indice)
    spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        DesdeCuenta.Text = rstCuenta!Cuenta
        HastaCuenta.Text = rstCuenta!Cuenta
        rstCuenta.Close
                Else
        DesdeCuenta.Text = WCuenta
        HastaCuenta.Text = WCuenta
    End If
    DesdeCuenta.SetFocus
    
End Sub

Sub Form_Load()
    Tipo1.Value = False
    Tipo2.Value = False
    Tipo3.Value = False
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeCuenta.Text = " "
    HastaCuenta.Text = "999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    TipoList.Clear
    
    TipoList.AddItem "Completo"
    TipoList.AddItem "Resumido"
    
    TipoList.ListIndex = 0
    
End Sub

