VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCcprv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Proveedores"
   ClientHeight    =   7770
   ClientLeft      =   1485
   ClientTop       =   750
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9120
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   8400
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   " "
      Top             =   3000
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2655
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox TipoFecha 
         Height          =   315
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   3
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   2
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wccprv.rpt"
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
      Left            =   -480
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   4155
      ItemData        =   "ccprv.frx":0000
      Left            =   480
      List            =   "ccprv.frx":0007
      TabIndex        =   6
      Top             =   3480
      Width           =   7695
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   8160
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   8160
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCcprv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Acumula As Double
Private Pasa As Single
Private WSaldo As Double
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String
Dim XSaldo As Double
Dim XTotal As Double
Dim XPagos As Double
Dim XDiferencia As Double

Private Sub Acepta_Click()

    If TipoFecha.ListIndex = 1 Then
        Tipo2.Value = True
        Tipo1.Value = False
    End If

    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores"
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
        .Index = "ClaveImpre"
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
    
    ZSaldo = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    If Tipo1.Value = True Then
        spCtaprv = "ListaCtaprvDesdeHastaSaldoTotal " + XParam
            Else
        spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    End If
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
        With RstCtaPrv
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WSaldo = !Saldo
                    Call Redondeo(WSaldo)
                    
                    If Tipo2.Value = True Or WSaldo <> 0 Then
                    
                        If TipoFecha.ListIndex = 0 Or (!OrdFecha >= WDesdeFecha And !OrdFecha <= WHastaFecha) Then
                        
                            XProveedor = !Proveedor
                            XLetra = !Letra
                            XTipo = !Tipo
                            XPunto = !Punto
                            XNumero = !Numero
                            XFecha = !Fecha
                            XEstado = !Estado
                            Xvencimiento = !Vencimiento
                            XVencimiento1 = !Vencimiento1
                            XNroInterno = !NroInterno
                            XTotal = !Total
                            XSaldo = !Saldo
                            XClave = !Clave
                            XOrdFecha = !OrdFecha
                            XOrdVencimiento = !OrdVencimiento
                            XImpre = !Impre
                            
                            With rstImpCtaCtePrv
                            
                                .Index = "CtaCte"
                                .Seek "=", XClave
                                If .NoMatch Then
                                    .AddNew
                                    !Proveedor = XProveedor
                                    !Letra = XLetra
                                    !Tipo = XTipo
                                    !Punto = XPunto
                                    !Numero = XNumero
                                    !Fecha = XFecha
                                    !Estado = XEstado
                                    !Vencimiento = Xvencimiento
                                    !Vencimiento1 = XVencimiento1
                                    !NroInterno = XNroInterno
                                    !Total = XTotal
                                    !Saldo = XSaldo
                                    !Clave = XClave
                                    !OrdFecha = XOrdFecha
                                    !OrdVencimiento = XOrdVencimiento
                                    !Impre = XImpre
                                    !Titulo = WTitulo
                                    If TipoFecha.ListIndex = 1 Then
                                        !Saldo = XTotal
                                    End If
                                    .Update
                                    .Bookmark = .LastModified
                                End If
                            End With
                        
                        End If
                        
                        If TipoFecha.ListIndex = 1 And !OrdFecha < WDesdeFecha Then
                            ZSaldo = ZSaldo + !Total
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
    
    If TipoFecha.ListIndex = 1 Then
    
        XLetra = ""
        XTipo = "00"
        XPunto = "0000"
        XNumero = "00000000"
        XFecha = "00/00/0000"
        XEstado = 0
        Xvencimiento = "00/00/0000"
        XVencimiento1 = "00/00/0000"
        XNroInterno = 0
        XTotal = ZSaldo
        XSaldo = ZSaldo
        XClave = ""
        XOrdFecha = "00000000"
        XOrdVencimiento = "00000000"
        XImpre = "SI"
        
        With rstImpCtaCtePrv
        
            .Index = "CtaCte"
            .Seek "=", XClave
            If .NoMatch Then
                .AddNew
                !Proveedor = XProveedor
                !Letra = XLetra
                !Tipo = XTipo
                !Punto = XPunto
                !Numero = XNumero
                !Fecha = XFecha
                !Estado = XEstado
                !Vencimiento = Xvencimiento
                !Vencimiento1 = XVencimiento1
                !NroInterno = XNroInterno
                !Total = XTotal
                !Saldo = XSaldo
                !Clave = XClave
                !OrdFecha = XOrdFecha
                !OrdVencimiento = XOrdVencimiento
                !Impre = XImpre
                !Titulo = WTitulo
                If TipoFecha.ListIndex = 1 Then
                    !Saldo = XTotal
                End If
                .Update
                .Bookmark = .LastModified
            End If
        End With
    
    End If
    
    
    Pasa = 0
    Acumula = 0
    
    da = ""
    With rstImpCtaCtePrv
        .Index = "ClaveImpre"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                If Pasa = 0 Then
                    Pasa = 1
                    Acumula = 0
                    Corte = !Proveedor
                    WProveedor = !Proveedor
                    WNombre = ""
                    WCheque = ""
                
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        WNombre = RstProveedor!Nombre
                        WCheque = RstProveedor!NombreCheque
                        RstProveedor.Close
                    End If
                End If
                If Corte <> !Proveedor Then
                    Acumula = 0
                    Corte = !Proveedor
                    WProveedor = !Proveedor
                    WNombre = ""
                    WCheque = ""
                
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        WNombre = RstProveedor!Nombre
                        WCheque = RstProveedor!NombreCheque
                        RstProveedor.Close
                    End If
                End If
                .Edit
                !SaldoList = 0
                If !Proveedor >= Desde.Text And !Proveedor <= Hasta.Text Then
                    WSaldo = !Saldo
                    Call Redondeo(WSaldo)
                    !SaldoList = WSaldo
                    Acumula = Acumula + WSaldo
                    !Acumulado = Acumula
                End If
                
                !Nombre = WNombre
                !Cheque = WCheque
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With

    If Tipo1.Value = True Then
        Listado.GroupSelectionFormula = "{CtaCtePrv.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {CtaCtePrv.Saldolist} <> 0.00"
            Else
        Listado.GroupSelectionFormula = "{CtaCtePrv.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {CtaCtePrv.Saldolist} <> 999999.99"
    End If
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = Wempresa + "Auxi.mdb"

    Listado.Action = 1
End Sub

Private Sub Cancela_Click()
    With rstEmpresa
        .Close
    End With
    With rstImpCtaCtePrv
        .Close
    End With
    Desde.SetFocus
    PrgCcprv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

    da = ""
    With rstRevisa
        .Index = "Ctacte"
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
    
    spCtaprv = "ListaCtaCtePrv"
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                XSaldo = !Saldo
                Call Redondeo(XSaldo)
                
                If XSaldo <> 0 Or Val(!Tipo) = 4 Then
                    
                    If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Or Val(!Tipo) = 4 Then
            
                        XProveedor = !Proveedor
                        XLetra = !Letra
                        XTipo = !Tipo
                        XPunto = !Punto
                        XNumero = !Numero
                        XFecha = !Fecha
                        XNroInterno = !NroInterno
                        XTotal = !Total
                        XSaldo = !Saldo
                        XClave = !Clave
                        XOrdFecha = !OrdFecha
                        
                        XSaldo = !Saldo
                        Call Redondeo(XSaldo)
                        XTotal = !Total
                        Call Redondeo(XTotal)
                        
                        With rstRevisa
                
                            .Index = "CtaCte"
                            .Seek "=", XClave
                            If .NoMatch Then
                                .AddNew
                                !Proveedor = XProveedor
                                !Letra = XLetra
                                !Tipo = XTipo
                                !Punto = XPunto
                                !Numero = XNumero
                                !Fecha = XFecha
                                !FechaOrd = XOrdFecha
                                !Importe = XTotal
                                !Saldo = XSaldo
                                !Saldorecalculado = 0
                                !Diferencia = 0
                                !NroInterno = XNroInterno
                                !Clave = XClave
                                .Update
                                .Bookmark = .LastModified
                            End If
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
    
    
    da = ""
    With rstRevisa
        .Index = "Ctacte"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                XPagos = 0
                
                XProveedor = !Proveedor
                XTipo = !Tipo
                XLetra = !Letra
                XPunto = !Punto
                XNumero = !Numero
                
                Sql1 = "Select Proveedor, Tipo1, Letra1, Punto1, Numero1, Importe1"
                Sql2 = " FROM Pagos"
                Sql3 = " Where Pagos.Proveedor = " + "'" + XProveedor + "'"
                Sql4 = " and Pagos.Tipo1 = " + "'" + XTipo + "'"
                Sql5 = " and Pagos.Letra1 = " + "'" + XLetra + "'"
                Sql6 = " and Pagos.Punto1 = " + "'" + XPunto + "'"
                Sql7 = " and Pagos.Numero1 = " + "'" + XNumero + "'"
                spPagos = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                If rstPagos.RecordCount > 0 Then
    
                    With rstPagos
                        .MoveFirst
                
                        If .NoMatch = False Then
                        Do
                
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            XPagos = XPagos + !Importe1
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                        End If
            
                    End With
        
                    rstPagos.Close
            
                End If
                
                Call Redondeo(XPagos)
                !Saldorecalculado = !Importe - XPagos
                XDiferencia = !Saldo - !Saldorecalculado
                Call Redondeo(XDiferencia)
                !Diferencia = XDiferencia
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
Stop

End Sub

Private Sub Command2_Click()

    spCtaprv = "ListaCtaCtePrv"
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Then
                
                    ZZClave = !Clave
                    ZZClaveiI = !Proveedor + !Letra + !Tipo + !Punto + !Numero
                    If ZZClave <> ZZClaveiI Then Stop
            
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


Stop
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
            
    Rem Pantalla.Visible = True
    Ayuda.Text = ""
    Rem AyudA.SetFocus

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Revisa
    OPEN_FILE_Empresa
    OPEN_FILE_ImpCtaCtePrv
End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Desde.Text = RstProveedor!Proveedor
        Hasta.Text = RstProveedor!Proveedor
        RstProveedor.Close
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    TipoFecha.Clear
    
    TipoFecha.AddItem "Toda la informacion"
    TipoFecha.AddItem "Entre fechas"
    
    TipoFecha.ListIndex = 0

    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "

    Desde.Text = "0"
    Hasta.Text = "99999999999"
    Panta.Value = False
    Impresora.Value = True
    Tipo1.Value = True
    Tipo2.Value = False
    Frame2.Visible = True
    Call Consulta_Click
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrdConsulta"
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

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            Desdefecha.SetFocus
        End If
    End If
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Desdefecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub

