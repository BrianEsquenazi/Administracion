VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSaldoCta 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Saldos de Cuentas Corrientes de Clientes"
   ClientHeight    =   6225
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   7200
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   5655
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   5055
         Begin VB.OptionButton Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   3360
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Documentos 
            Caption         =   "Documentos"
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton CtaCte 
            Caption         =   "Cuenta Corriente"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Moneda"
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
         Begin VB.OptionButton Dolares 
            Caption         =   "Dolares"
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Pesos 
            Caption         =   "Pesos"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   285
            Width           =   1215
         End
      End
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6480
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wsaldocta.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Clientes"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "SaldoCta.frx":0000
      Left            =   120
      List            =   "SaldoCta.frx":0007
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgSaldoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    spCtacte = "ModificaCtacteTipo1"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    spCtacte = "ModificaCtacteTipo2"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    spCtacte = "ModificaCtacteImporte0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    If CtaCte.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte1 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte2 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte3 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte4 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Documentos.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte5 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte6 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte7 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte8 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Total.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte9 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte10 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte11 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte12 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    If CtaCte.Value = True Then
        WTitulo = "Cuenta Corriente - "
    End If
    If Documentos.Value = True Then
        WTitulo = "Documentos - "
    End If
    If Total.Value = True Then
        WTitulo = "Total - "
    End If
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Saldos de Cuenta Corriente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{CtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.SelectionFormula = "{CtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CtaCte.Cliente, CtaCte.Importe3, Cliente.Razon " _
                        + "From " + DSQ + ".dbo.CtaCte CtaCte, " _
                        + DSQ + ".dbo.Cliente Cliente " _
                        + "WHERE " _
                        + "CtaCte.Cliente = Cliente.Cliente AND " _
                        + "CtaCte.Cliente >= '" + Desde.Text + "' AND " _
                        + "CtaCte.Cliente <= '" + Hasta.Text + "' AND " _
                        + "CtaCte.Importe3 <> 0."
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgSaldoCta.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()
    spCtacte = "BorrarCtacte " + "'" + "040001441301" + "'"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + "     " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
        End With
    End If
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Desde.Text = rstCliente!Cliente
        Hasta.Text = rstCliente!Cliente
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub
Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Pesos.Value = True
    Dolares.Value = False
    CtaCte.Value = True
    Documentos.Value = False
    Total.Value = False
    Frame2.Visible = True
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    With rstClientes
        .Index = "Razon"
        .MoveFirst
        Do
            If .EOF = False Then
            
                DA = Len(!Razon) - WEspacios
                
                For aa = 1 To DA
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                        Auxi = !Cliente
                        IngresaItem = Auxi + "    " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    End If

End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

