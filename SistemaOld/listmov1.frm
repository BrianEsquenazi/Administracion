VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListmov1 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Movimientos de Envases por Cliente"
   ClientHeight    =   6375
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6375
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   2400
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hastafecha 
         Height          =   300
         Left            =   1440
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   1440
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox HastaCli 
         Height          =   300
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox DesdeCli 
         Height          =   300
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wmovenv1.rpt"
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "listmov1.frx":0000
      Left            =   120
      List            =   "listmov1.frx":0007
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListmov1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim XParam As String
Dim Vector(10000, 6) As String

Private Sub Acepta_Click()

    DesdeCli.Text = UCase(DesdeCli.Text)
    HastaCli.Text = UCase(HastaCli.Text)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    Da = 0
    With rstFichaenv
        .Index = "Envase"
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
    
    Inicial = 0
    
    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    Pasa = 0
    
    XParam = "'" + DesdeCli.Text + "','" _
                + HastaCli.Text + "'"
    spMovenv = "ListaMovenvDesdeHastaCliente" + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
   
        With rstMovenv
    
            .MoveFirst
            
            Do
            
                    If Pasa = 0 Then
                        Pasa = 1
                        Corte = rstMovenv!Cliente
                        Corte1 = rstMovenv!Envase
                    End If
                    
                    If Corte <> rstMovenv!Cliente Or Corte1 <> rstMovenv!Envase Then
                    
                        If Inicial <> 0 Then
                    
                            With rstFichaenv
                                .AddNew
                                !Cliente = WCliente
                                !Envase = WEnvase
                                !Fecha = "00/00/0000"
                                !FechaOrd = "00000000"
                                !Tipo = 0
                                !Numero = 0
                                !Inicial = Inicial
                                !Observaciones = "Saldo Inicial"
                                !Lista1 = "Mov.Var."
                                !Lista2 = ""
                                !Titulo = WTitulo
                                .Update
                            End With
                        End If
                        
                        Corte = rstMovenv!Cliente
                        Corte1 = rstMovenv!Envase
                        Inicial = 0
                        
                    End If

                    WCodigo = rstMovenv!Codigo
                    WEnvase = rstMovenv!Envase
                    WCantidad = rstMovenv!Cantidad
                    WFecha = rstMovenv!Fecha
                    WFechaord = rstMovenv!FechaOrd
                    WEnvase = rstMovenv!Envase
                    WMovi = rstMovenv!Movimiento
                    WCliente = rstMovenv!Cliente

                    If WFechaord < WDesde Then
                        If WMovi = "E" Then
                            Inicial = Inicial - WCantidad
                                Else
                            Inicial = Inicial + WCantidad
                        End If
                    End If
                    
                    If WFechaord >= WDesde And WFechaord <= WHasta Then
                
                        With rstFichaenv
                            .AddNew
                            !Cliente = WCliente
                            !Envase = WEnvase
                            !Fecha = WFecha
                            !FechaOrd = WFechaord
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = 0
                                !Salida = WCantidad
                                    Else
                                !Entrada = WCantidad
                                !Salida = 0
                            End If
                            !Observaciones = ""
                            !Lista1 = "Mov.Var."
                            !Lista2 = ""
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
    End If
    
    If Inicial <> 0 Then
                    
        With rstFichaenv
            .AddNew
            !Cliente = WCliente
            !Envase = WEnvase
            !Fecha = "00/00/0000"
            !FechaOrd = "00000000"
            !Tipo = 0
            !Numero = 0
            !Inicial = Inicial
            !Observaciones = "Saldo Inicial"
            !Lista1 = "Mov.Var."
            !Lista2 = ""
            !Titulo = WTitulo
            .Update
        End With
        
    End If
    
    Da = 0
    With rstFichaenv
        .Index = "Envase"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WDescriEnvases = ""
                WDescriCliente = ""
                WEnvase = Str$(!Envase)
                WCliente = !Cliente
                
                spEnvases = "ConsultaEnvases " + "'" + WEnvase + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvases.RecordCount > 0 Then
                    WDescriEnvase = rstEnvases!Descripcion
                End If
                
                spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WDescriCliente = rstCliente!Razon
                End If
                
                !DescriEnvases = WDescriEnvase
                !DescriCliente = WDescriCliente
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.WindowTitle = "Listado de MOvimiento de Envases por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{FichaEnv.Cliente} in " + Chr$(34) + DesdeCli.Text + Chr$(34) + " to " + Chr$(34) + HastaCli.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "WMovEnv1.rpt"
            Else
        Listado.ReportFileName = "WMovEnv1Resu.rpt"
    End If
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstFichaenv
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    DesdeCli.SetFocus
    PrgListmov1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desdecli_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCli.Text = UCase(DesdeCli.Text)
        HastaCli.Text = DesdeCli.Text
        HastaCli.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaEnv
End Sub

Private Sub Hastacli_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCli.Text = UCase(HastaCli.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCli.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListmov1.Caption = "Listado de Movimientos de Envases por Cliente :  " + !Nombre
        End If
    End With
    DesdeCli.Text = ""
    HastaCli.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaCliente"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCliente
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                Pantalla.AddItem IngresaItem
                IngresaItem = rstCliente!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstCliente.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesdeCli.Text = rstCliente!Cliente
        HastaCli.Text = rstCliente!Cliente
    End If
    
    DesdeCli.SetFocus

    
End Sub

