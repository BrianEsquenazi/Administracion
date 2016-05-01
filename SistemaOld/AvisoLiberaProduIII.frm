VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoLiberaProduIII 
   AutoRedraw      =   -1  'True
   Caption         =   "Aviso de Entrada de Productos para Devolucion"
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
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Aviso de Productos a Etiquetar como PT o DY"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   5535
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
Attribute VB_Name = "PrgAvisoLiberaProduIII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim WVector(1000) As String
Dim LeeAviso(100, 3) As String

Dim LugarVector As Integer

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    WEmpresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    LugarVector = 0

    Sql1 = "Select *"
    Sql2 = " FROM LiberaTerminado"
    Sql3 = " Where LiberaTerminado.ImpreProdI = " + "'" + "S" + "'"
    Sql4 = " and LiberaTerminado.TipoPro = 'FA'"
    Sql5 = ""

    spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstLiberaTerminado.RecordCount > 0 Then
        LugarVector = 1
        rstLiberaTerminado.Close
    End If
    
    If LugarVector > 0 Then
        PrgAvisoLiberaProduIII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgAvisoLiberaProduIII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Listado.ReportFileName = ""

    Listado.WindowTitle = "Listado de Pedidos Pendientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.ReportFileName = "ListaLiberadoI.RPT"
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT LiberaTerminado.Producto, LiberaTerminado.Fecha, LiberaTerminado.Partida, LiberaTerminado.PartiOri, LiberaTerminado.Observa, LiberaTerminado.ImpreProdI, LiberaTerminado.ImpreProdII, LiberaTerminado.ImpreProdIII " _
                + "From " _
                + DSQ + ".dbo.LiberaTerminado LiberaTerminado " _
                + "WHERE " _
                + "LiberaTerminado.ImpreProdI = 'S' AND " _
                + "LiberaTerminado.TipoPro >= 'FA' AND " _
                + "LiberaTerminado.TipoPro <= 'FA'"
                
    Listado.Connect = Connect()

    Listado.Action = 1
        
    Sql1 = "UPDATE LiberaTerminado SET "
    Sql2 = " ImpreProdI = " + "'" + "" + "'"
    Sql3 = " Where LiberaTerminado.ImpreProdI = " + "'" + "S" + "'"
    Sql4 = " and LiberaTerminado.TipoPro = 'FA'"
    Sql5 = ""
    
    spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgAvisoLiberaProduIII.Hide
    Unload Me
    Close
    End
End Sub



