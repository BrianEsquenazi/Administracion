VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPosdat 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cheque Posdatados"
   ClientHeight    =   5205
   ClientLeft      =   2205
   ClientTop       =   1935
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2535
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox HastaBanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   16
         Text            =   " "
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Desdebanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   15
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wposdat.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "posdat.frx":0000
      Left            =   840
      List            =   "posdat.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPosdat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo Control_Error
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    da = 0
    With rstPosdat
        .Index = "Impre"
        .Seek ">=", 0
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
                If .EOF = True Then
                    Exit Do
                End If
            
                Rem If rstPagos!TipoOrd <> 5 Then
            
                If rstPagos!Banco2 >= Val(DesdeBanco.Text) And rstPagos!Banco2 <= Val(HastaBanco.Text) Then
                    
                    WFechaCheque = Right$(rstPagos!Fecha2, 4) + Mid$(rstPagos!Fecha2, 4, 2) + Left$(rstPagos!Fecha2, 2)
                    
                    If !Fecha2 <> !Fecha Then
                    
                        If WFechaCheque >= WDesde And WFechaCheque <= WHasta Then
                    
                            WBanco = rstPagos!Banco2
                            WFecha = rstPagos!Fecha
                            WImporte = rstPagos!Importe2
                            WCheque = rstPagos!Numero2
                            Wvencimiento = rstPagos!Fecha2
                            WProveedor = rstPagos!Proveedor
                            WObservaciones = rstPagos!Observaciones
                
                            With rstPosdat
                                .AddNew
                                !Banco = WBanco
                                !Fecha = WFecha
                                !Cheque = WCheque
                                !Proveedor = WProveedor
                                !Importe = WImporte
                                !Vencimiento = Wvencimiento
                                !FechaOrd = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                                !Observaciones = Left$(WObservaciones, 30)
                                !Titulo = WTitulo
                                !Titulo1 = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
                                .Update
                            End With
                        End If
                    End If
                End If
                
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstPagos.Close
    End If
    
    da = 0
    With rstPosdat
        .Index = "Impre"
        .Seek ">=", 0
        If .NoMatch = False Then
            Do
                .Edit
                
                WObservaciones = !Observaciones
                WNombreBanco = ""
                
                WProveedor = !Proveedor
                WBanco = !Banco
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WObservaciones = Left$(RstProveedor!Nombre, 20)
                    RstProveedor.Close
                End If
                
                spBanco = "ConsultaBanco " + "'" + Str$(WBanco) + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    WNombreBanco = Left$(rstBanco!Nombre, 20)
                    rstBanco.Close
                End If
                
                !Observaciones = WObservaciones
                !NombreBanco = WNombreBanco
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Posdat = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado Cheque Posdatados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Posdat.Banco} in " + Desde.Text + " to " + Hasta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    Exit Sub
    
Control_Error:
     coderr = Err
     Resume Next
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgPosdat.Hide
    Unload Me
    Menu.Show
End Sub



Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Posdat
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeBanco.SetFocus
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.SetFocus
    End If
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeBanco.Text = "0"
    HastaBanco.Text = "9999"
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
        Desde.Text = rstBanco!Banco
        Hasta.Text = rstBanco!Banco
        rstBanco.Close
                Else
        Desde.Text = WBanco
        Hasta.Text = WBanco
    End If
    Desde.SetFocus
    
End Sub


