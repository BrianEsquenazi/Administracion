VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListmov2 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Movimientos de Envases por Cliente"
   ClientHeight    =   6000
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8025
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   8025
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2775
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
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
      Begin VB.TextBox HastaEnv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox DesdeEnv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         MaxLength       =   4
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
         Caption         =   "Hasta Envase"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Envase"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wmovenv2.rpt"
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
      Height          =   2700
      ItemData        =   "listmov2.frx":0000
      Left            =   120
      List            =   "listmov2.frx":0007
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListmov2"
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
    Pasa = 0
    
    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
            
    XParam = "'" + DesdeEnv.Text + "','" _
                + HastaEnv.Text + "'"
    spMovenv = "ListaMovenvDesdeHastaEnvases" + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
    
            .MoveFirst
            
            Do
            
                    If Pasa = 0 Then
                        Pasa = 1
                        Corte = rstMovenv!Envase
                    End If
                    
                    If Corte <> rstMovenv!Envase Then
                    
                        If Inicial <> 0 Then
                    
                            With rstFichaenv
                                .AddNew
                                !Cliente = ""
                                !Envase = Corte
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
                        
                        Corte = rstMovenv!Envase
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
            !Cliente = ""
            !Envase = Corte
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
    
    Listado.WindowTitle = "Listado de MOvimiento de Envases por Envases"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{FichaEnv.Envase} in " + DesdeEnv.Text + " to " + HastaEnv.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
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
    
    DesdeEnv.SetFocus
    PrgListmov2.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeEnv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaEnv.Text = DesdeEnv.Text
        HastaEnv.SetFocus
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

Private Sub HastaEnv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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
        DesdeEnv.SetFocus
    End If
End Sub



Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListmov2.Caption = "Listado de movimientos de Envaese por Envases :  " + !Nombre
        End If
    End With
    DesdeEnv.Text = " "
    HastaEnv.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spEnvases = "ListaEnvases"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEnvases
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstEnvases!Envases
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstEnvases.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spEnvases = "ConsultaEnvases " + "'" + Claveven$ + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        DesdeEnv.Text = rstEnvases!Envases
        HastaEnv.Text = rstEnvases!Envases
    End If
    
    DesdeEnv.SetFocus

    
End Sub


