VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProyPrv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyeccion de Cobros"
   ClientHeight    =   6330
   ClientLeft      =   3165
   ClientTop       =   1305
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   6135
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence2 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   1
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Parametros de Fechas"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WProyPrv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Proveedores"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "proyprv.frx":0000
      Left            =   600
      List            =   "proyprv.frx":0007
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgProyPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String

Private Sub Acepta_Click()

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Auxi1 = Vence1.Text
            !Auxi2 = Vence2.Text
            !Auxi3 = Vence3.Text
            !Auxi4 = Vence4.Text
            .Update
        End If
    End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    Listado.WindowTitle = "Listado de Proyeccion de Pagos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)

    da = ""
    With RstProve
        .Index = "Proveedor"
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
            Do
                If !Saldo <> 0 Then
                
                    WSaldo = !Saldo
                    Wvencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    WProveedor = !Proveedor
                    
                    With RstProve
                        .Index = "Proveedor"
                        .Seek "=", WProveedor
                        If .NoMatch = False Then
                            .Edit
                                Else
                            .AddNew
                            !Proveedor = WProveedor
                        End If
                        !Importe6 = !Importe6 + WSaldo
                        If Wvencimiento <= Fecha1 Then
                            !Importe1 = !Importe1 + WSaldo
                                Else
                            If Wvencimiento <= Fecha2 Then
                                !Importe2 = !Importe2 + WSaldo
                                    Else
                                If Wvencimiento <= Fecha3 Then
                                    !Importe3 = !Importe3 + WSaldo
                                        Else
                                    If Wvencimiento <= Fecha4 Then
                                        !Importe4 = !Importe4 + WSaldo
                                            Else
                                        !Importe5 = !Importe5 + WSaldo
                                    End If
                                End If
                            End If
                        End If
                        !Auxi1 = Vence1.Text
                        !Auxi2 = Vence2.Text
                        !Auxi3 = Vence3.Text
                        !Auxi4 = Vence4.Text
                        !Titulo = WTitulo
                        .Update
                    End With
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    RstCtaPrv.Close
    End If
    
    
    da = ""
    With RstProve
        .Index = "Proveedor"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WProveedor = !Proveedor
                WNombre = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    RstProveedor.Close
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

    Listado.GroupSelectionFormula = "{Proveedor.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
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
    With RstProve
        .Close
    End With
    Desde.SetFocus
    PrgProyPrv.Hide
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
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Prove
    OPEN_FILE_Auxiliar
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
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
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vence1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Sub Form_Load()
    Desde.Text = "0"
    Hasta.Text = "99999999999"
    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Vence1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence2.SetFocus
                Else
            Vence1.SetFocus
        End If
    End If
End Sub

Private Sub Vence2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence3.SetFocus
                Else
            Vence2.SetFocus
        End If
    End If
End Sub

Private Sub Vence3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence4.SetFocus
                Else
            Vence3.SetFocus
        End If
    End If
End Sub

Private Sub Vence4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
End Sub

