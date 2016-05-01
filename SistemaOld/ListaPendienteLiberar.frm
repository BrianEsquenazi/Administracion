VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPendienteLiberar 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Productos Pendientes de Liberar"
   ClientHeight    =   3165
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   4815
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
         Left            =   2640
         TabIndex        =   5
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
         Left            =   840
         TabIndex        =   4
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
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   600
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
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaPendienteLiberar.rpt"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaPendienteLiberar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Tres As String
Dim XParam As String

Private Sub Acepta_Click()

    WMarca = ""

    Sql1 = "UPDATE EntDev SET "
    Sql2 = " Trabajo = " + "'" + WMarca + "'"
    spEntdev = Sql1 + Sql2
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)

    WMarca = "X"

    Sql1 = "UPDATE EntDev SET "
    Sql2 = " Trabajo = " + "'" + WMarca + "'"
    Sql3 = " Where Saldo > 0"
    spEntdev = Sql1 + Sql2 + Sql3
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Listado de Productos Pendientes de Liberar"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "Entdev.Codigo, Entdev.Fecha, Entdev.Terminado, Entdev.Cantidad, Entdev.Observaciones, Entdev.Lote, Entdev.Cliente, Entdev.Saldo, Entdev.PartiOri, Entdev.ImpreTerminado, Entdev.Trabajo, Entdev.Estado, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Entdev Entdev, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Entdev.Cliente = Cliente.Cliente AND " _
                    + "Entdev.Trabajo = 'X' AND " _
                    + "Entdev.Estado = 'NK'"
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaPendienteLiberar.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

