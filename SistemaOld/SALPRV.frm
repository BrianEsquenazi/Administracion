VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSalprv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Saldos de Cuentas Corrientes de Proveedores"
   ClientHeight    =   4485
   ClientLeft      =   2625
   ClientTop       =   2235
   ClientWidth     =   6780
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   6780
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   1
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
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
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5280
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wsaldoprv.rpt"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2595
      ItemData        =   "SALPRV.frx":0000
      Left            =   120
      List            =   "SALPRV.frx":0007
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgSalprv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String
Dim Vector(1000, 2) As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Saldos Cuenta Corriente de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    da = ""
    With rstImpCtaCtePrv
        .Index = "ClaveImpre"
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHastaSaldo " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                XProveedor = !Proveedor
                XSaldo = !Saldo
                XClave = !Clave
                
                If !Saldo <> 0 Then
                
                With rstImpCtaCtePrv
                
                    .Index = "CtaCte"
                    .Seek "=", XClave
                    If .NoMatch Then
                        .AddNew
                        !Proveedor = XProveedor
                        !Saldo = XSaldo
                        !Clave = XClave
                        !Titulo = WTitulo
                        !Tipo = ""
                        !Numero = ""
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstCtaPrv.Close
    
    End If
    
    Erase Vector
    Renglon = 0
    
    With rstImpCtaCtePrv
            .Index = "CtaCte"
            .MoveFirst
            Do
                .Edit
                !SaldoList = 0
                If Val(!Proveedor) >= Val(Desde.Text) And Val(!Proveedor) <= Val(Hasta.Text) Then
                    !SaldoList = !Saldo
                End If
                
                WProveedor = !Proveedor
                WNombre = ""
                WCheque = ""
                
                Pasa = "S"
                
                For Ciclo = 1 To Renglon
                    If Vector(Ciclo, 1) = WProveedor Then
                        WNombre = Vector(Ciclo, 2)
                        Pasa = "N"
                        Exit For
                    End If
                Next Ciclo
                    
                If Pasa = "S" Then
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                         WNombre = RstProveedor!Nombre
                         WCheque = RstProveedor!NombreCheque
                         RstProveedor.Close
                    End If
                    !Nombre = WNombre
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = WProveedor
                    Vector(Renglon, 2) = WNombre
                End If
                
                !Nombre = WNombre
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With

    Listado.GroupSelectionFormula = "{CtaCteprv.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
   
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstImpCtaCtePrv
        .Close
    End With
    
    Desde.SetFocus
    PrgSalprv.Hide
    Unload Me
    Menu.Show
    
End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImpCtaCtePrv
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Sub Form_Load()
    Desde.Text = "0"
    Hasta.Text = "99999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
            
    Pantalla.Visible = True
    Rem Ayuda.Text = ""
    Rem AyudA.SetFocus

End Sub


Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Desde.Text = RstProveedor!Proveedor
        Hasta.Text = RstProveedor!Proveedor
        RstProveedor.Close
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub

