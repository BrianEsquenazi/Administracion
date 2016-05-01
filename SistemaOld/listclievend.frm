VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListClieVend 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Clientes por Vendedor"
   ClientHeight    =   2910
   ClientLeft      =   2160
   ClientTop       =   1095
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Hasta 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   8
         Text            =   " "
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Desde 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
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
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Vendedor"
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
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Vendedor"
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
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListClieVend.rpt"
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
End
Attribute VB_Name = "PrgListClieVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            Nomempresa = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Clientes por Vendedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Cliente.Vendedor} in " + Desde.Text + " to " + Hasta.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.Direccion, Cliente.Localidad, Cliente.Provincia, Cliente.Postal, Cliente.Vendedor, " _
                    + "Vendedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Vendedor Vendedor " _
                    + "Where " _
                    + "Cliente.Vendedor = Vendedor.Vendedor AND " _
                    + "Cliente.Vendedor >= " + Desde.Text + " AND " _
                    + "Cliente.Vendedor <= " + Hasta.Text
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgListClieVend.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "0"
    Hasta.Text = "9999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub




