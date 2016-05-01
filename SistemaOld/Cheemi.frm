VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCheEmi 
   Caption         =   "Listado de Cheques Emitidos"
   ClientHeight    =   4830
   ClientLeft      =   3615
   ClientTop       =   2010
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4830
   ScaleWidth      =   5655
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4680
      TabIndex        =   16
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox HastaBanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   " "
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox DesdeBanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Text            =   " "
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   282
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCheemi.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCheEmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial() As Variant ' Matriz de 2 dimensiones que contiene registros
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim XParam As String

Private Sub Acepta_Click()
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
            .Update
        End If
    End With

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    da = 0
    With rstMovban
        .Index = "Clave"
        .Seek "=", da
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
    
    spPagos = "ListaPagos"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
    With rstPagos
            .MoveFirst
            Do

                If !FechaOrd > "19991017" Then
                
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                    If Val(!Tipo2) = 2 Then
                        If Val(!Banco2) >= Val(DesdeBanco) And Val(!Banco2) <= Val(HastaBanco) Then
                            WBanco = !Banco2
                            WOrden = !Orden
                            WFecha = !Fecha
                            WFechaord = !FechaOrd
                            WAcredita = !Fecha2
                            WAcreditaOrd = !FechaOrd2
                            WObservaciones = ""
                            WObservaciones = !Observaciones
                            Rem If Val(!Proveedor) = 0 Then
                            Rem      WObservaciones = !Observaciones
                            Rem         Else
                            Rem     With rstProveedor
                            Rem         .Index = "Proveedor"
                            Rem         .Seek "=", !Proveedor
                            Rem        If .NoMatch = False Then
                            Rem             WObservaciones = !Nombre
                            Rem         End If
                            Rem     End With
                            Rem End If
                            WNumero = !Numero2
                            WImporte = !Importe2
                            WOrden = !Orden
                            WProveedor = !Proveedor
                
                            With rstMovban
                                .AddNew
                                !da = 0
                                !Banco = WBanco
                                !Fecha = WFecha
                                !FechaOrd = WFechaord
                                !Acredita = WAcredita
                                !AcreditaOrd = WAcreditaOrd
                                !Observaciones = WObservaciones
                                !Numero = WNumero
                                !Debito = 0
                                !Credito = WImporte
                                !Comprobante = WOrden
                                !Empresa = 1
                                !Titulo = WTitulo
                                !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                !Proveedor = WProveedor
                                .Update
                            End With
                        End If
                    End If
                    
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    rstPagos.Close
    
    End If

    da = 0
    With rstMovban
        .Index = "Clave"
        .Seek "=", da
        If .NoMatch = False Then
            Do
                .Edit
                
                WBanco = !Banco
                WNombre = ""
                
                spBanco = "ConsultaBancos " + "'" + Str$(WBanco) + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    WNombre = rstBanco!Nombre
                    rstBanco.Close
                End If
                
                If Val(!Proveedor) <> 0 Then
                    spProveedor = "ConsultaProveedores " + "'" + !Proveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        WObservaciones = RstProveedor!Nombre
                        RstProveedor.Close
                    End If
                    !Observaciones = WObservaciones
                End If
                
                !Nombre = WNombre
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.GroupSelectionFormula = "{Movban.banco} in " + DesdeBanco + " to " + HastaBanco
    Rem Listado.GroupSelectionFormula = "{Movban.banco} in 0 to 9999"
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstMovban
        .Close
    End With
    Desde.SetFocus
    PrgCheEmi.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Movban
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeBanco.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.Text = DesdeBanco.Text
        HastaBanco.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeBanco.Text = 0
    HastaBanco.Text = 9999
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
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

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WBanco = WIndice.List(Indice)
    spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        DesdeBanco.Text = rstBanco!Banco
        HastaBanco.Text = rstBanco!Banco
        rstBanco.Close
                Else
        DesdeBanco.Text = WBanco
        HastaBanco.Text = WBanco
    End If
    Desde.SetFocus
    
End Sub

