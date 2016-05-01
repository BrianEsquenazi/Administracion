VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCcprvFec 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Proveedores a Fecha"
   ClientHeight    =   7770
   ClientLeft      =   1485
   ClientTop       =   750
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9120
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   " "
      Top             =   3600
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3135
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   5175
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   840
         TabIndex        =   15
         Top             =   2160
         Width           =   3375
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1560
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   4
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   3
         Text            =   " "
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Emision"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wccprvfec.rpt"
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
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   3375
      ItemData        =   "CcprvFec.frx":0000
      Left            =   480
      List            =   "CcprvFec.frx":0007
      TabIndex        =   7
      Top             =   4080
      Width           =   7695
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7560
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7560
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCcprvFec"
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
Dim rstPagos As Recordset
Dim spPagos As String
Dim cParam As String
Dim XParam As String
Dim WProveedor As String
Dim WTipo As String
Dim WLetra As String
Dim WPunto As String
Dim WNumero As String
Dim WImporte As Double
Dim WClave As String

Private Sub Acepta_Click()

    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + WDia

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
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
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Proveedor <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.OrdFecha <= " + "'" + WFecha + "'"
    spCtaprv = ZSql
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    Rem XParam = "'" + Desde.Text + "','" _
    rem              + Hasta.Text + "'"
    Rem spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Rem Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    Rem If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If !OrdFecha <= WFecha Then
            
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
                        !Fecha1 = Fecha.Text
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
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
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Pagos.Proveedor <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and Pagos.FechaOrd > " + "'" + WFecha + "'"
    ZSql = ZSql + " and Pagos.Importe1 <> 0"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
    
    Rem spPagos = "ListaPagos"
    Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPagos.RecordCount > 0 Then
            
        With rstPagos
            .MoveFirst
            Do
                        
                If WFecha < !FechaOrd Then
                
                If Val(!Proveedor) >= Val(Desde.Text) And Val(!Proveedor) <= Val(Hasta.Text) Then
                
                If !Importe1 <> 0 Then
                
                    WProveedor = !Proveedor
                    WTipo = !Tipo1
                    WLetra = !Letra1
                    WPunto = !Punto1
                    WNumero = !Numero1
                    WImporte = !Importe1
                    
                    Call Ceros(WProveedor, 11)
                    Call Ceros(WTipo, 2)
                    Call Ceros(WPunto, 4)
                    Call Ceros(WNumero, 8)
                    
                    WClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                           
                    With rstImpCtaCtePrv
                        .Index = "Ctacte"
                        .Seek "=", WClave
                        If .NoMatch = False Then
                            .Edit
                            !Saldo = !Saldo + WImporte
                            .Update
                        End If
                    End With
                    
                End If
                
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstPagos.Close
        
        
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM AplicaProve"
    ZSql = ZSql + " Where AplicaProve.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and AplicaProve.Proveedor <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and AplicaProve.OrdFecha > " + "'" + WFecha + "'"
    spAplicaProve = ZSql
    Set RstAplicaProve = db.OpenRecordset(spAplicaProve, dbOpenSnapshot, dbSQLPassThrough)
    If RstAplicaProve.RecordCount > 0 Then
    
    Rem spAplicaProve = "ListaAplicaProve"
    Rem Set rstAplicaProve = db.OpenRecordset(spAplicaProve, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstAplicaProve.RecordCount > 0 Then
            
        With RstAplicaProve
            .MoveFirst
            Do
                        
                WProveedor = !Proveedor
                WTipo = !Tipo
                WLetra = !Letra
                WPunto = !Punto
                WNumero = !Numero
                WImporte = !Importe
                
                Call Ceros(WProveedor, 11)
                Call Ceros(WTipo, 2)
                Call Ceros(WPunto, 4)
                Call Ceros(WNumero, 8)
                
                WClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                       
                With rstImpCtaCtePrv
                    .Index = "Ctacte"
                    .Seek "=", WClave
                    If .NoMatch = False Then
                        .Edit
                        !Saldo = !Saldo + WImporte
                        .Update
                    End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        RstAplicaProve.Close
        
        
    End If
    
    
    
    
    
    Pasa = 0
    Acumula = 0
    Corte = ""

    With rstImpCtaCtePrv
            .Index = "ClaveImpre"
            .MoveFirst
            Do
            
                If Pasa = 0 Then
                    Pasa = 1
                    Acumula = 0
                    Corte = !Proveedor
 Rem nan
                      WProveedor = !Proveedor
                
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
   Rem nan
                     WProveedor = !Proveedor
                
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
                
                WProveedor = !Proveedor
         Rem       WNombre = ""
                WCheque = ""
                
             Rem   If Corte <> WProveedor Then
             Rem       spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
             Rem       Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
             Rem       If RstProveedor.RecordCount > 0 Then
             Rem           WNombre = RstProveedor!Nombre
             Rem           WCheque = RstProveedor!NombreCheque
             Rem           RstProveedor.Close
             Rem       End If
             Rem       Corte = WProveedor
             Rem   End If
                
                !Nombre = WNombre
                !Cheque = WCheque
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
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
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"

    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstImpCtaCtePrv
        .Close
    End With
    Desde.SetFocus
    PrgCcprvFec.Hide
    Unload Me
    Menu.Show
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

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
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
        Fecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
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


