VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaGeneral 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado General de Productos"
   ClientHeight    =   3060
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   4815
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   1
         Text            =   " "
         Top             =   840
         Width           =   735
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   735
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
         Left            =   2400
         TabIndex        =   6
         Top             =   1320
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
         Left            =   960
         TabIndex        =   5
         Top             =   1320
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
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Linea"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Linea"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaGeneral.rpt"
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
Attribute VB_Name = "PrgListaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Acepta_Click()

    Rem On Error GoTo WError
    
    ZBlanco = ""
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase1  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase2  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase3  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase4  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase6  = " + "'" + ZBlanco + "',"
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZBlanco + "'"
                     
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    ZMarca = "X"
    
    Rem
    Rem Bolsas x 20
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase1  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 13 "
    ZSql = ZSql + " or Envase2 = 13 "
    ZSql = ZSql + " or Envase3 = 13 "
    ZSql = ZSql + " or Envase4 = 13 "
    ZSql = ZSql + " or Envase5 = 13 "
    ZSql = ZSql + " or Envase6 = 13 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem Bolsas x 25
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase2  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 9 "
    ZSql = ZSql + " or Envase2 = 9 "
    ZSql = ZSql + " or Envase3 = 9 "
    ZSql = ZSql + " or Envase4 = 9 "
    ZSql = ZSql + " or Envase5 = 9 "
    ZSql = ZSql + " or Envase6 = 9 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem Tambores x 120
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase3  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 12 "
    ZSql = ZSql + " or Envase2 = 12 "
    ZSql = ZSql + " or Envase3 = 12 "
    ZSql = ZSql + " or Envase4 = 12 "
    ZSql = ZSql + " or Envase5 = 12 "
    ZSql = ZSql + " or Envase6 = 12 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase3  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 21 "
    ZSql = ZSql + " or Envase2 = 21 "
    ZSql = ZSql + " or Envase3 = 21 "
    ZSql = ZSql + " or Envase4 = 21 "
    ZSql = ZSql + " or Envase5 = 21 "
    ZSql = ZSql + " or Envase6 = 21 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem Tambores x 150
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase4  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 52 "
    ZSql = ZSql + " or Envase2 = 52 "
    ZSql = ZSql + " or Envase3 = 52 "
    ZSql = ZSql + " or Envase4 = 52 "
    ZSql = ZSql + " or Envase5 = 52 "
    ZSql = ZSql + " or Envase6 = 52 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem Tambores x 200
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 2 "
    ZSql = ZSql + " or Envase2 = 2 "
    ZSql = ZSql + " or Envase3 = 2 "
    ZSql = ZSql + " or Envase4 = 2 "
    ZSql = ZSql + " or Envase5 = 2 "
    ZSql = ZSql + " or Envase6 = 2 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 5 "
    ZSql = ZSql + " or Envase2 = 5 "
    ZSql = ZSql + " or Envase3 = 5 "
    ZSql = ZSql + " or Envase4 = 5 "
    ZSql = ZSql + " or Envase5 = 5 "
    ZSql = ZSql + " or Envase6 = 5 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 20 "
    ZSql = ZSql + " or Envase2 = 20 "
    ZSql = ZSql + " or Envase3 = 20 "
    ZSql = ZSql + " or Envase4 = 20 "
    ZSql = ZSql + " or Envase5 = 20 "
    ZSql = ZSql + " or Envase6 = 20 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 23 "
    ZSql = ZSql + " or Envase2 = 23 "
    ZSql = ZSql + " or Envase3 = 23 "
    ZSql = ZSql + " or Envase4 = 23 "
    ZSql = ZSql + " or Envase5 = 23 "
    ZSql = ZSql + " or Envase6 = 23 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase5  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 25 "
    ZSql = ZSql + " or Envase2 = 25 "
    ZSql = ZSql + " or Envase3 = 25 "
    ZSql = ZSql + " or Envase4 = 25 "
    ZSql = ZSql + " or Envase5 = 25 "
    ZSql = ZSql + " or Envase6 = 25 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem Cist. x 1000
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase6  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 11 "
    ZSql = ZSql + " or Envase2 = 11 "
    ZSql = ZSql + " or Envase3 = 11 "
    ZSql = ZSql + " or Envase4 = 11 "
    ZSql = ZSql + " or Envase5 = 11 "
    ZSql = ZSql + " or Envase6 = 11 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem
    Rem pallets
    Rem
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 69 "
    ZSql = ZSql + " or Envase2 = 69 "
    ZSql = ZSql + " or Envase3 = 69 "
    ZSql = ZSql + " or Envase4 = 69 "
    ZSql = ZSql + " or Envase5 = 69 "
    ZSql = ZSql + " or Envase6 = 69 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 70 "
    ZSql = ZSql + " or Envase2 = 70 "
    ZSql = ZSql + " or Envase3 = 70 "
    ZSql = ZSql + " or Envase4 = 70 "
    ZSql = ZSql + " or Envase5 = 70 "
    ZSql = ZSql + " or Envase6 = 70 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 71 "
    ZSql = ZSql + " or Envase2 = 71 "
    ZSql = ZSql + " or Envase3 = 71 "
    ZSql = ZSql + " or Envase4 = 71 "
    ZSql = ZSql + " or Envase5 = 71 "
    ZSql = ZSql + " or Envase6 = 71 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 72 "
    ZSql = ZSql + " or Envase2 = 72 "
    ZSql = ZSql + " or Envase3 = 72 "
    ZSql = ZSql + " or Envase4 = 72 "
    ZSql = ZSql + " or Envase5 = 72 "
    ZSql = ZSql + " or Envase6 = 72 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 73 "
    ZSql = ZSql + " or Envase2 = 73 "
    ZSql = ZSql + " or Envase3 = 73 "
    ZSql = ZSql + " or Envase4 = 73 "
    ZSql = ZSql + " or Envase5 = 73 "
    ZSql = ZSql + " or Envase6 = 73 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 74 "
    ZSql = ZSql + " or Envase2 = 74 "
    ZSql = ZSql + " or Envase3 = 74 "
    ZSql = ZSql + " or Envase4 = 74 "
    ZSql = ZSql + " or Envase5 = 74 "
    ZSql = ZSql + " or Envase6 = 74 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Terminado SET "
    ZSql = ZSql + " ImpreEnvase7  = " + "'" + ZMarca + "'"
    ZSql = ZSql + " Where Envase1 = 75 "
    ZSql = ZSql + " or Envase2 = 75 "
    ZSql = ZSql + " or Envase3 = 75 "
    ZSql = ZSql + " or Envase4 = 75 "
    ZSql = ZSql + " or Envase5 = 75 "
    ZSql = ZSql + " or Envase6 = 75 "
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado General de Productos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Terminado.Linea} in " + Desde.Text + " to " + Hasta.Text
    Dos = " and {Terminado.ListaProducto} = 1 "
    Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.Linea, Terminado.Clase, Terminado.Carga, Terminado.EstadoProducto, Terminado.ListaProducto, Terminado.ImpreEnvase1, Terminado.ImpreEnvase2, Terminado.ImpreEnvase3, Terminado.ImpreEnvase4, Terminado.ImpreEnvase5, Terminado.ImpreEnvase6, Terminado.ImpreEnvase7, " _
                + "Lineas.Nombre " _
                + "From " _
                + DSQ + ".dbo.Terminado Terminado, " _
                + DSQ + ".dbo.Lineas Lineas " _
                + "Where " _
                + "Terminado.Linea = Lineas.Linea AND " _
                + "Terminado.Linea >= " + Desde.Text + " AND " _
                + "Terminado.Linea <= " + Hasta.Text + " AND " _
                + "Terminado.ListaProducto = 1"
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaGeneral.Hide
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
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


