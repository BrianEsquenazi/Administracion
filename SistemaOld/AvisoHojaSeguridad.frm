VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoHojaSeguridad 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Control de Hojas de Seguridad"
   ClientHeight    =   4605
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3975
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Mensaje1 
         Alignment       =   2  'Center
         Caption         =   "Hay Productos Terminados con Hojas Tecnicas que hay que ratificar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaAvisoHojaSeguridad.rpt"
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
Attribute VB_Name = "PrgAvisoHojaSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WEntra = "N"
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Terminado"
    ZSql = ZSql & " Where Terminado.EstadoHoja = " + "'" + "N" + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WEntra = "S"
        rstTerminado.Close
    End If
    
    
    If WEntra = "S" Then
        PrgAvisoHojaSeguridad.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgAvisoHojaSeguridad.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Listado.WindowTitle = "Listado de Pedidos Pendientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.Destination = 1
    Rem Listado.Destination = 0

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.EstadoHoja " _
            + "From " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "Terminado.EstadoHoja = 'N'"

    Listado.Connect = Connect()
    Listado.Action = 1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgAvisoHojaSeguridad.Hide
    Unload Me
    Close
    End
End Sub


