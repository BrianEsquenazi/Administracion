VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProyec 
   AutoRedraw      =   -1  'True
   Caption         =   "Proyeccion de Entrada de Materias Primas"
   ClientHeight    =   6330
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   6840
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4335
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   960
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Parametros de Fechas"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WProyec.rpt"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "proyec.frx":0000
      Left            =   600
      List            =   "proyec.frx":0007
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgProyec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XOrden As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    With rstProyec
        .Index = "Producto"
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

    With rstAuxiliar
        .MoveFirst
        If .NoMatch = False Then
            .Edit
            !Auxi1 = Vence1.Text
            !Auxi2 = Vence2.Text
            !Auxi3 = Vence3.Text
            .Update
        End If
    End With
        
    Listado.WindowTitle = "Proyeccion de Entradas de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)

    spOrden = "ListaOrdenTotal"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstOrden.RecordCount > 0 Then
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
                WOrden = rstOrden!Orden
                WArticulo = rstOrden!Articulo
                WFecha = Right$(rstOrden!fecha2, 4) + Mid$(rstOrden!fecha2, 4, 2) + Left$(rstOrden!fecha2, 2)
                WCantidad = rstOrden!Cantidad
                WRecibida = rstOrden!Recibida
                WSaldo = rstOrden!Cantidad - rstOrden!Recibida
                
                If WSaldo > 0 Then
                    
                    With rstProyec
                        .AddNew
                        !Orden = WOrden
                        !Producto = WArticulo
                        !Canti5 = !Canti5 + WSaldo
                        If WFecha <= Fecha1 Then
                            !Canti1 = !Canti1 + WSaldo
                                Else
                            If WFecha <= fecha2 Then
                                !Canti2 = !Canti2 + WSaldo
                                    Else
                                If WFecha <= Fecha3 Then
                                    !Canti3 = !Canti3 + WSaldo
                                        Else
                                    !Canti4 = !Canti4 + WSaldo
                                End If
                            End If
                        End If
                        XOrden = Str$(WOrden)
                        Call Ceros(XOrden, 6)
                        !Clave = !Producto + XOrden
                        .Update
                    End With
                    
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    End If
    
    With rstProyec
        .Index = "Producto"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WProducto = !Producto
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WProducto + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                End If
                !Descripcion = WDescripcion
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
                
    Listado.GroupSelectionFormula = "{Proyec.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstProyec
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgProyec.Hide
    Unload Me
    Menu.Show
    
End Sub



Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Auxiliar
    OPEN_FILE_Empresa
    OPEN_FILE_Proyec
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgProyec.Caption = "Proyecion de Entrada de Materia Prima :  " + !Nombre
        End If
    End With
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

