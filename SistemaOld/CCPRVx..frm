VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   3735
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
      Left            =   1200
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wccprvx.rpt"
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
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   4155
      ItemData        =   "CCPRVx..frx":0000
      Left            =   480
      List            =   "CCPRVx..frx":0007
      TabIndex        =   6
      Top             =   3480
      Width           =   7695
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6360
      TabIndex        =   1
      Top             =   840
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

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
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
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstCtaPrv.Close
    
    End If
    
    
    Pasa = 0
    Acumula = 0

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
                    Corte = !Proveedor
                End If
                If Corte <> !Proveedor Then
                    Acumula = 0
                    Corte = !Proveedor
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
                WNombre = ""
                WCheque = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    WCheque = RstProveedor!NombreCheque
                    RstProveedor.Close
                End If
                
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
    PrgCcprv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrd"
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
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                    
                    
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


