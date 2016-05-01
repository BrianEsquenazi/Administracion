VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListdepo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Depositos"
   ClientHeight    =   4785
   ClientLeft      =   2925
   ClientTop       =   2415
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4935
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   1
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Banco"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Banco"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wdepositos.rpt"
      Destination     =   1
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
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "listdepo.frx":0000
      Left            =   0
      List            =   "listdepo.frx":0007
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListdepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim XParam As String

Private Sub Acepta_Click()
    
    Listado.WindowTitle = "Listado de Depositos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    spDeposito = "ModificaDepositoImpolista0"
    Set rstDeposito = db.OpenRecordset(spDeposito, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "','" _
                 + WDesde + "','" _
                 + WHasta + "'"
    spDeposito = "ModificaDepositoImpolista" + XParam
    Set rstDeposito = db.OpenRecordset(spDeposito, dbOpenSnapshot, dbSQLPassThrough)

    Rem With rstDepositos
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             .Edit
    Rem             !IMPOListA = 0
    Rem             If !Banco >= Val(Desde.Text) And !Banco <= Val(Hasta.Text) Then
    Rem                 If !FechaOrd >= WDesde And !FechaOrd <= WHasta Then
    Rem                     !IMPOListA = !Importe2
    Rem                 End If
    Rem             End If
    Rem             .Update
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With

    Uno = "{Depositos.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Depositos.Banco} in " + Desde.Text + " to " + Hasta.Text
    Listado.GroupSelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Depositos.Deposito, Depositos.Banco, Depositos.Fecha, Depositos.FechaOrd, Depositos.Tipo2, Depositos.Numero2, Depositos.Fecha2, Depositos.Observaciones2, Depositos.Impolista, Banco.Nombre " _
                        + "From " + DSQ + ".dbo.Depositos Depositos, " _
                        + DSQ + ".dbo.Banco Banco " _
                        + "Where Depositos.Banco = Banco.Banco AND Depositos.Banco >= " + Desde.Text + " AND Depositos.Banco <= " + Hasta.Text + " AND Depositos.FechaOrd >= '" + WDesde + "' AND Depositos.FechaOrd <= '" + WHasta + "'"
    Listado.DataFiles(2) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    Desde.SetFocus
    PrgListdepo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = Str$(rstBanco!Banco)
                    Call Ceros(Auxi, 4)
                    IngresaItem = Auxi + " " + rstBanco!Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstBanco!Banco
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
            
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WBanco = WIndice.List(Indice)
    spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        Desde.Text = rstBanco!Banco
        Hasta.Text = rstBanco!Banco
        rstBanco.Close
                Else
        Desde.Text = WBanco
        Hasta.Text = WBanco
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
        Desdefecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

