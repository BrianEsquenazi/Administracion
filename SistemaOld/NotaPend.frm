VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgNotaPend 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Notas de Debito por Diferencia de Cambio Pendoientes"
   ClientHeight    =   3825
   ClientLeft      =   2400
   ClientTop       =   2070
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   7170
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WNotaPend.rpt"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "NotaPend.frx":0000
      Left            =   840
      List            =   "NotaPend.frx":0007
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgNotaPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    Rem spCtacte = "ModificaCtacteImporteIva0"
    Rem Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    Rem spCtacte = "ModificaCtacteIva4"
    Rem Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Importe4 = 0"
    ZSql = ZSql + " Where Impre = " + "'" + "ND" + "'"
    ZSql = ZSql + " or Impre = " + "'" + "NC" + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Importe4 = Saldo"
    ZSql = ZSql + " Where Impre = " + "'" + "ND" + "'"
    ZSql = ZSql + " or Impre = " + "'" + "NC" + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Notas de Debito por Diferencia de Cambio"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.Vencimiento, CtaCte.Saldo, CtaCte.SaldoUs, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Paridad, CtaCte.Importe4, " _
                    + "Cliente.Razon, Cliente.Pago1, " _
                    + "Pago.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Pago Pago " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "Cliente.Pago1 = Pago.Pago AND " _
                    + "CtaCte.Tipo >= '04' AND CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '00000000' AND CtaCte.OrdFecha <= '99999999' AND " _
                    + "CtaCte.Importe4 <> 0."
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    PrgNotaPend.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


