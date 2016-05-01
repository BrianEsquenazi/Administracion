VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEmiteCertificado 
   Caption         =   "Emision de Ceritificado de Analisis"
   ClientHeight    =   5910
   ClientLeft      =   735
   ClientTop       =   645
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   10425
   Begin VB.ListBox PantallaII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "emitecertificado.frx":0000
      Left            =   240
      List            =   "emitecertificado.frx":0007
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox Idioma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   26
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox ZZClave 
         Height          =   285
         Left            =   5280
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox ZZEmpresa 
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8520
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Limpia 
         Caption         =   "  Limpia Pantalla"
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
         Left            =   8520
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Descripcion 
         Enabled         =   0   'False
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
         Left            =   2280
         LinkTimeout     =   100
         MaxLength       =   30
         TabIndex        =   16
         Text            =   " "
         Top             =   1920
         Width           =   5775
      End
      Begin VB.TextBox Cantidad 
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
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   13
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "  "
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   327680
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox Cliente 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   1
         Text            =   " "
         Top             =   960
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
         Left            =   8520
         TabIndex        =   7
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
         Height          =   495
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Idioma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   27
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label DesProducto 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label DesCliente 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Producto Terminado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9840
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "eti1.rpt"
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
      Left            =   8760
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "emitecertificado.frx":0015
      Left            =   240
      List            =   "emitecertificado.frx":001C
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7560
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEmiteCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstCertificado As Recordset
Dim spCertificado As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String

Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 6) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String
Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String

Dim WTerminado(100, 3) As String
Dim LugarTerminado As Integer

Private CargaEmpresa(12, 2) As String

Private Sub Cancela_click()

    PrgEmiteCertificado.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    spClientes = "ListaClienteConsulta"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then

    With rstClientes
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstClientes!Cliente + " " + rstClientes!razon
                Pantalla.AddItem IngresaItem
                IngresaItem = rstClientes!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstClientes.Close

    End If
    
    PantallaII.Visible = False
    Pantalla.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
            
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = Claveven$
        DesCliente.Caption = rstCliente!razon
        rstCliente.Close
    End If
    
    ZClave = Cliente.Text + Terminado.Text
    spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Descripcion.Text = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
        rstPrecios.Close
    End If
    
    Pantalla.Visible = False
    Ayuda.Visible = False
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        WEspacios = Len(Ayuda.Text)
        WIndice.Clear
    
        Pantalla.Clear
            
        spCliente = "ListaClienteConsulta"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        Da = Len(rstCliente!razon) - WEspacios
                        
                        For aa = 1 To Da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!razon, aa, WEspacios) Then
                                Auxi = rstCliente!Cliente
                                IngresaItem = Auxi + "    " + rstCliente!razon
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCliente!Cliente
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
            rstCliente.Close
        End If
        
    End If

End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Impresora"
    Tipo.AddItem "Pantalla"
    Tipo.AddItem "Pantalla Word"
    
    Tipo.ListIndex = 0
    
    Idioma.Clear
    
    Idioma.AddItem "Castellano"
    Idioma.AddItem "Ingles"
    
    Idioma.ListIndex = 0

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    DesProducto.Caption = ""
    Cantidad.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
End Sub

Sub Limpia_Click()

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    DesProducto.Caption = ""
    Cantidad.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Lote.SetFocus
    
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        XEmpresa = Wempresa
    
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
                CargaEmpresa(6, 1) = "0010"
                CargaEmpresa(6, 2) = "Empresa10"
                CargaEmpresa(7, 1) = "0011"
                CargaEmpresa(7, 2) = "Empresa11"
                ZHasta = 7
            Case Else
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
                ZHasta = 4
        End Select
        
        Entra = "N"
        LugarTerminado = 0
        Erase WTerminado
            
        For ZCiclo = 1 To ZHasta
            
            Wempresa = CargaEmpresa(ZCiclo, 1)
            txtOdbc = CargaEmpresa(ZCiclo, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Prueter"
            ZSql = ZSql + " Where Prueter.Lote = " + "'" + Lote.Text + "'"
            rsPrueter = ZSql
            Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueter.RecordCount > 0 Then
                LugarTerminado = LugarTerminado + 1
                WTerminado(LugarTerminado, 1) = rstPrueter!Producto
                WTerminado(LugarTerminado, 2) = rstPrueter!Prueba
                WTerminado(LugarTerminado, 3) = Wempresa
                rstPrueter.Close
            End If
            
        Next ZCiclo
        
        Call Conecta_Empresa
        
        If LugarTerminado > 0 Then
        
            If LugarTerminado = 1 Then
            
                Terminado.Text = WTerminado(1, 1)
                ZZClave.Text = WTerminado(1, 2)
                ZZEmpresa.Text = WTerminado(1, 3)
            
                XEmpresa = Wempresa
                Select Case Val(XEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2, 4, 8, 9
                        Wempresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
            
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Terminado"
                ZSql = ZSql & " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    DesProducto.Caption = rstTerminado!Descripcion
                    Descripcion.Text = rstTerminado!Descripcion
                    Cliente.SetFocus
                End If
                
                Call Conecta_Empresa
                
                    Else
                    
                Dim IngresaItem As String
                PantallaII.Clear
    
                For CiclaTerminado = 1 To LugarTerminado
                    IngresaItem = WTerminado(CiclaTerminado, 1)
                    PantallaII.AddItem IngresaItem
                Next CiclaTerminado
                
                PantallaII.Visible = True
                Pantalla.Visible = False
                
            End If
            
        End If
        
    End If
    
End Sub

Private Sub PantallaII_Click()

    Indice = PantallaII.ListIndex + 1
    
    Terminado.Text = WTerminado(Indice, 1)
    ZZClave.Text = WTerminado(Indice, 2)
    ZZEmpresa.Text = WTerminado(Indice, 3)
    
    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Terminado"
    ZSql = ZSql & " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesProducto.Caption = rstTerminado!Descripcion
        Descripcion.Text = rstTerminado!Descripcion
        Cliente.SetFocus
    End If

    Call Conecta_Empresa
    
End Sub


Private Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
        
            XEmpresa = Wempresa
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2, 4, 8, 9
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                DesCliente.Caption = rstClientes!razon
                rstClientes.Close
                
                spPrecios = "ConsultaPrecios " + "'" + Cliente.Text + Terminado.Text + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Descripcion.Text = Trim(Left$(rstPrecios!Descripcion, 50))
                    rstPrecios.Close
                        Else
                    Descripcion.Text = DesProducto.Caption
                End If
            End If
            
            Call Conecta_Empresa
            
            Cantidad.SetFocus
            
                Else
                
            Descripcion.Text = DesProducto.Caption
            Cantidad.SetFocus
            
        End If
    End If
End Sub

Private Sub Acepta_Click()
    
    XEmpresa = Wempresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            ZZZCliente = "S00102"
        Case 2, 4, 8, 9
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            ZZZCliente = "P99999"
        Case Else
    End Select
    
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        WIdioma = IIf(IsNull(rstClientes!Idioma), "0", rstClientes!Idioma)
        rstClientes.Close
    End If
    
    Call Conecta_Empresa
    
    If WIdioma = 1 Then
        Idioma.ListIndex = 1
    End If


    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    
    
    ZArticulo = Terminado.Text
    ZProducto = Terminado.Text
    ZLote = Lote.Text
    ZCantidad = Cantidad.Text
    ZCliente = Cliente.Text
        
    Erase ZOpcion
    Erase ZValor
    Erase ZEnsayo
    Erase ZStd
    Erase ZDescri
    Erase ZDescriII
        
    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    ZVersion = 0
    
    ZZEntra = "N"
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM AltaCertificado"
    ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
    ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
    spAltaCertificado = ZSql
    Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
    If rstAltaCertificado.RecordCount > 0 Then
        ZOpcion(1) = rstAltaCertificado!Opcion1
        ZOpcion(2) = rstAltaCertificado!Opcion2
        ZOpcion(3) = rstAltaCertificado!Opcion3
        ZOpcion(4) = rstAltaCertificado!Opcion4
        ZOpcion(5) = rstAltaCertificado!Opcion5
        ZOpcion(6) = rstAltaCertificado!Opcion6
        ZOpcion(7) = rstAltaCertificado!Opcion7
        ZOpcion(8) = rstAltaCertificado!Opcion8
        ZOpcion(9) = rstAltaCertificado!Opcion9
        ZOpcion(10) = rstAltaCertificado!Opcion10
        rstAltaCertificado.Close
        ZZEntra = "S"
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM AltaCertificado"
        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
        ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZZZCliente + "'"
        spAltaCertificado = ZSql
        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
        If rstAltaCertificado.RecordCount > 0 Then
            ZOpcion(1) = rstAltaCertificado!Opcion1
            ZOpcion(2) = rstAltaCertificado!Opcion2
            ZOpcion(3) = rstAltaCertificado!Opcion3
            ZOpcion(4) = rstAltaCertificado!Opcion4
            ZOpcion(5) = rstAltaCertificado!Opcion5
            ZOpcion(6) = rstAltaCertificado!Opcion6
            ZOpcion(7) = rstAltaCertificado!Opcion7
            ZOpcion(8) = rstAltaCertificado!Opcion8
            ZOpcion(9) = rstAltaCertificado!Opcion9
            ZOpcion(10) = rstAltaCertificado!Opcion10
            rstAltaCertificado.Close
            ZZEntra = "S"
        End If
    End If
            
    If ZZEntra = "N" Then
        m$ = "No esta definido el certificado de analisis para este producto"
        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If
            
    If ZZEntra = "S" Then
    
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
                CargaEmpresa(6, 1) = "0010"
                CargaEmpresa(6, 2) = "Empresa10"
                CargaEmpresa(7, 1) = "0011"
                CargaEmpresa(7, 2) = "Empresa11"
                ZHasta = 7
            Case Else
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
                ZHasta = 4
        End Select
            
        For ZCiclo = 1 To ZHasta
        
            If Val(CargaEmpresa(ZCiclo, 1)) = Val(ZZEmpresa.Text) Then
            
                Wempresa = CargaEmpresa(ZCiclo, 1)
                txtOdbc = CargaEmpresa(ZCiclo, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Prueter"
                ZSql = ZSql + " Where Prueter.Lote = " + "'" + ZLote + "'"
                rsPrueter = ZSql
                Set rstPrueter = db.OpenRecordset(rsPrueter, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrueter.RecordCount > 0 Then
                        
                    WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
                            
                    ZValor(1) = rstPrueter!Valor1
                    ZValor(2) = rstPrueter!valor2
                    ZValor(3) = rstPrueter!Valor3
                    ZValor(4) = rstPrueter!valor4
                    ZValor(5) = rstPrueter!valor5
                    ZValor(6) = rstPrueter!valor6
                    ZValor(7) = rstPrueter!valor7
                    ZValor(8) = rstPrueter!valor8
                    ZValor(9) = rstPrueter!valor9
                    ZValor(10) = rstPrueter!valor10
                        
                    rstPrueter.Close
                    
                    WFechaElaboracion = ""
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Lote.Text + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                        ZZHoja = rstHoja!Hoja
                        ZZProducto = rstHoja!Producto
                        zzrevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                        zzmesesrevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                        
                        If ZZFechaRevalida <> "  /  /    " And ZZFechaRevalida <> "00/00/0000" Then
                            WFecha = ZZFechaRevalida
                            WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        End If
                        
                        ZZFecha = rstHoja!Fecha
                        ZZMeses = ""
                        rstHoja.Close
                        
                        spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                            rstTerminado.Close
                        End If
                        
                        If Val(ZZMeses) <> 0 Then
                        
                            If Val(zzrevalida) <> 0 Then
                                WVida = Val(zzmesesrevalida)
                                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                WAno = Val(Right$(ZZFechaRevalida, 4))
                                    Else
                                WVida = Val(ZZMeses)
                                WMes = Val(Mid$(ZZFecha, 4, 2))
                                WAno = Val(Right$(ZZFecha, 4))
                            End If
                            
                            For Ciclo = 1 To WVida
                                WMes = WMes + 1
                                If WMes > 12 Then
                                    WAno = WAno + 1
                                    WMes = 1
                                End If
                            Next Ciclo
                            ZMes = Str$(WMes)
                            ZAno = Str$(WAno)
                            Call Ceros(ZMes, 2)
                            Call Ceros(ZAno, 4)
                            WFechaElaboracion = ZMes + "/" + ZAno
                            
                        End If
                        
                    End If
                        
                    If Left$(ZArticulo, 2) = "DW" Then
                        WProducto = "DW" + Mid$(ZArticulo, 3, 10)
                            Else
                        If Left$(ZArticulo, 2) = "SE" Then
                            WProducto = "SE" + Mid$(ZArticulo, 3, 10)
                                Else
                            WProducto = "PT" + Mid$(ZArticulo, 3, 10)
                        End If
                    End If
                        
                    Select Case Val(Wempresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            Wempresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            Wempresa = "0004"
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                        
                    LlamaImprime = "N"
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM EspecifUnificaVersion"
                    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
                    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                    spEspecifUnificaVersion = ZSql
                    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspecifUnificaVersion.RecordCount > 0 Then
                        With rstEspecifUnificaVersion
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                
                                    dada = rstEspecifUnificaVersion!Version
                                    
                                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                                            
                                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                                            
                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                
                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                
                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                
                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                
                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta2), "", rstEspecifUnificaVersion!Hasta2)
                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                        
                                        ZStd(1, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor1Ing), "", rstEspecifUnificaVersion!Valor1Ing)
                                        ZStd(2, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor2Ing), "", rstEspecifUnificaVersion!Valor2Ing)
                                        ZStd(3, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor3Ing), "", rstEspecifUnificaVersion!Valor3Ing)
                                        ZStd(4, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor4Ing), "", rstEspecifUnificaVersion!Valor4Ing)
                                        ZStd(5, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor5Ing), "", rstEspecifUnificaVersion!Valor5Ing)
                                        ZStd(6, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor6Ing), "", rstEspecifUnificaVersion!Valor6Ing)
                                        ZStd(7, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor7Ing), "", rstEspecifUnificaVersion!Valor7Ing)
                                        ZStd(8, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor8Ing), "", rstEspecifUnificaVersion!Valor8Ing)
                                        ZStd(9, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor9Ing), "", rstEspecifUnificaVersion!Valor9Ing)
                                        ZStd(10, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor10Ing), "", rstEspecifUnificaVersion!Valor10Ing)
                                                            
                                        ZStd(1, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor11Ing), "", rstEspecifUnificaVersion!Valor11Ing)
                                        ZStd(2, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor22Ing), "", rstEspecifUnificaVersion!Valor22Ing)
                                        ZStd(3, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor33Ing), "", rstEspecifUnificaVersion!Valor33Ing)
                                        ZStd(4, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor44Ing), "", rstEspecifUnificaVersion!Valor44Ing)
                                        ZStd(5, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor55Ing), "", rstEspecifUnificaVersion!Valor55Ing)
                                        ZStd(6, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor66Ing), "", rstEspecifUnificaVersion!Valor66Ing)
                                        ZStd(7, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor77Ing), "", rstEspecifUnificaVersion!Valor77Ing)
                                        ZStd(8, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor88Ing), "", rstEspecifUnificaVersion!Valor88Ing)
                                        ZStd(9, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor99Ing), "", rstEspecifUnificaVersion!Valor99Ing)
                                        ZStd(10, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010Ing), "", rstEspecifUnificaVersion!Valor1010Ing)
                                                
                                        ZVersion = rstEspecifUnificaVersion!Version
                                        LlamaImprime = "S"
                                        
                                        m$ = "ATENCION : La partida esta asociada a una version de especificaciones que no es la actual" + Chr$(13) + _
                                             "Version " + Str$(rstEspecifUnificaVersion!Version) + Chr$(13) + _
                                             "Fecha de vigencia del : " + rstEspecifUnificaVersion!FechaInicio + " al " + rstEspecifUnificaVersion!FechaFinal
                                        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                                
                                    End If
                                        
                                    If WDesde > WFechaord And LlamaImprime = "N" Then
                                            
                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                
                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                
                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                
                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                
                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta2), "", rstEspecifUnificaVersion!Hasta2)
                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                        
                                        ZStd(1, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor1Ing), "", rstEspecifUnificaVersion!Valor1Ing)
                                        ZStd(2, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor2Ing), "", rstEspecifUnificaVersion!Valor2Ing)
                                        ZStd(3, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor3Ing), "", rstEspecifUnificaVersion!Valor3Ing)
                                        ZStd(4, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor4Ing), "", rstEspecifUnificaVersion!Valor4Ing)
                                        ZStd(5, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor5Ing), "", rstEspecifUnificaVersion!Valor5Ing)
                                        ZStd(6, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor6Ing), "", rstEspecifUnificaVersion!Valor6Ing)
                                        ZStd(7, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor7Ing), "", rstEspecifUnificaVersion!Valor7Ing)
                                        ZStd(8, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor8Ing), "", rstEspecifUnificaVersion!Valor8Ing)
                                        ZStd(9, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor9Ing), "", rstEspecifUnificaVersion!Valor9Ing)
                                        ZStd(10, 5) = IIf(IsNull(rstEspecifUnificaVersion!Valor10Ing), "", rstEspecifUnificaVersion!Valor10Ing)
                                                            
                                        ZStd(1, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor11Ing), "", rstEspecifUnificaVersion!Valor11Ing)
                                        ZStd(2, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor22Ing), "", rstEspecifUnificaVersion!Valor22Ing)
                                        ZStd(3, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor33Ing), "", rstEspecifUnificaVersion!Valor33Ing)
                                        ZStd(4, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor44Ing), "", rstEspecifUnificaVersion!Valor44Ing)
                                        ZStd(5, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor55Ing), "", rstEspecifUnificaVersion!Valor55Ing)
                                        ZStd(6, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor66Ing), "", rstEspecifUnificaVersion!Valor66Ing)
                                        ZStd(7, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor77Ing), "", rstEspecifUnificaVersion!Valor77Ing)
                                        ZStd(8, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor88Ing), "", rstEspecifUnificaVersion!Valor88Ing)
                                        ZStd(9, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor99Ing), "", rstEspecifUnificaVersion!Valor99Ing)
                                        ZStd(10, 6) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010Ing), "", rstEspecifUnificaVersion!Valor1010Ing)
                                        
                                                
                                        ZVersion = rstEspecifUnificaVersion!Version
                                        LlamaImprime = "S"
                                        
                                        m$ = "ATENCION : La partida esta asociada a una version de especificaciones que no es la actual" + Chr$(13) + _
                                             "Version " + Str$(rstEspecifUnificaVersion!Version) + Chr$(13) + _
                                             "Fecha de vigencia del : " + rstEspecifUnificaVersion!FechaInicio + " al " + rstEspecifUnificaVersion!FechaFinal
                                        a% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                        
                                        
                                    End If
                                    
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstEspecifUnificaVersion.Close
                    End If
                    
                    Rem by nan 16-8-2013
                    If LlamaImprime = "N" Then
                        
                        Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,ensayo4,ensayo5,Ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
                        Sql2 = " FROM EspecifUnifica"
                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEspecifUnifica.RecordCount > 0 Then
                                
                            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
                            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
                            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
                            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
                            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
                            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
                            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
                            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
                            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
                            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
                                                
                            ZStd(1, 1) = rstEspecifUnifica!Valor1
                            ZStd(2, 1) = rstEspecifUnifica!valor2
                            ZStd(3, 1) = rstEspecifUnifica!Valor3
                            ZStd(4, 1) = rstEspecifUnifica!valor4
                            ZStd(5, 1) = rstEspecifUnifica!valor5
                            ZStd(6, 1) = rstEspecifUnifica!valor6
                            ZStd(7, 1) = rstEspecifUnifica!valor7
                            ZStd(8, 1) = rstEspecifUnifica!valor8
                            ZStd(9, 1) = rstEspecifUnifica!valor9
                            ZStd(10, 1) = rstEspecifUnifica!valor10
                                                
                            ZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                            ZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                            ZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                            ZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                            ZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                            ZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                            ZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                            ZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                            ZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                            ZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                            rstEspecifUnifica.Close
                        End If
                            
                        Sql1 = "Select desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Valor1Ing,Valor2Ing,Valor3Ing,Valor4Ing,Valor5Ing,Valor6Ing,Valor7Ing,Valor8Ing,Valor9Ing,Valor10Ing,Valor11Ing,Valor22Ing,Valor33Ing,Valor44Ing,Valor55Ing,Valor66Ing,Valor77Ing,Valor88Ing,Valor99Ing,Valor1010Ing,Version"
                        Sql2 = " FROM EspecifUnifica"
                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEspecifUnifica.RecordCount > 0 Then
                            
                            
                            ZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                            ZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                            ZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                            ZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                            ZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                            ZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                            ZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                            ZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                            ZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                            ZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                    
                            ZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                            ZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
                            ZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                            ZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                            ZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                            ZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                            ZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                            ZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                            ZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                            ZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                                
                            ZStd(1, 5) = IIf(IsNull(rstEspecifUnifica!Valor1Ing), "", rstEspecifUnifica!Valor1Ing)
                            ZStd(2, 5) = IIf(IsNull(rstEspecifUnifica!Valor2Ing), "", rstEspecifUnifica!Valor2Ing)
                            ZStd(3, 5) = IIf(IsNull(rstEspecifUnifica!Valor3Ing), "", rstEspecifUnifica!Valor3Ing)
                            ZStd(4, 5) = IIf(IsNull(rstEspecifUnifica!Valor4Ing), "", rstEspecifUnifica!Valor4Ing)
                            ZStd(5, 5) = IIf(IsNull(rstEspecifUnifica!Valor5Ing), "", rstEspecifUnifica!Valor5Ing)
                            ZStd(6, 5) = IIf(IsNull(rstEspecifUnifica!Valor6Ing), "", rstEspecifUnifica!Valor6Ing)
                            ZStd(7, 5) = IIf(IsNull(rstEspecifUnifica!Valor7Ing), "", rstEspecifUnifica!Valor7Ing)
                            ZStd(8, 5) = IIf(IsNull(rstEspecifUnifica!Valor8Ing), "", rstEspecifUnifica!Valor8Ing)
                            ZStd(9, 5) = IIf(IsNull(rstEspecifUnifica!Valor9Ing), "", rstEspecifUnifica!Valor9Ing)
                            ZStd(10, 5) = IIf(IsNull(rstEspecifUnifica!Valor10Ing), "", rstEspecifUnifica!Valor10Ing)
                                                
                            ZStd(1, 6) = IIf(IsNull(rstEspecifUnifica!Valor11Ing), "", rstEspecifUnifica!Valor11Ing)
                            ZStd(2, 6) = IIf(IsNull(rstEspecifUnifica!Valor22Ing), "", rstEspecifUnifica!Valor22Ing)
                            ZStd(3, 6) = IIf(IsNull(rstEspecifUnifica!Valor33Ing), "", rstEspecifUnifica!Valor33Ing)
                            ZStd(4, 6) = IIf(IsNull(rstEspecifUnifica!Valor44Ing), "", rstEspecifUnifica!Valor44Ing)
                            ZStd(5, 6) = IIf(IsNull(rstEspecifUnifica!Valor55Ing), "", rstEspecifUnifica!Valor55Ing)
                            ZStd(6, 6) = IIf(IsNull(rstEspecifUnifica!Valor66Ing), "", rstEspecifUnifica!Valor66Ing)
                            ZStd(7, 6) = IIf(IsNull(rstEspecifUnifica!Valor77Ing), "", rstEspecifUnifica!Valor77Ing)
                            ZStd(8, 6) = IIf(IsNull(rstEspecifUnifica!Valor88Ing), "", rstEspecifUnifica!Valor88Ing)
                            ZStd(9, 6) = IIf(IsNull(rstEspecifUnifica!Valor99Ing), "", rstEspecifUnifica!Valor99Ing)
                            ZStd(10, 6) = IIf(IsNull(rstEspecifUnifica!Valor1010Ing), "", rstEspecifUnifica!Valor1010Ing)
                                                
                            ZVersion = rstEspecifUnifica!Version
                            
                            
                            rstEspecifUnifica.Close
                            LlamaImprime = "S"
                        End If
                    
                    End If
                    
                    If LlamaImprime = "S" Then
                        
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(1) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(1) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(2) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(2) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(3) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(3) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(4) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(4) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(5) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(5) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(6) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(6) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(7) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(7) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(8) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(8) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(9) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(9) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
            
                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnsayo.RecordCount > 0 Then
                            If Idioma.ListIndex = 0 Then
                                ZDescri(10) = rstEnsayo!Descripcion
                                    Else
                                ZDescri(10) = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                            End If
                            ZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                            rstEnsayo.Close
                        End If
                                
                        Call Conecta_Empresa
                        
                        XEmpresa = Wempresa
                        Select Case Val(XEmpresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                Wempresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2, 4, 8, 9
                                Wempresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZRazon = ""
                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCliente.RecordCount > 0 Then
                            ZRazon = Left$(rstCliente!razon, 50)
                            ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                            rstCliente.Close
                        End If
                        
                        Rem by nan ***** 22-5-2015 se debe imprimir siempre fecha de vencimiento ***********
                        Rem If ZImpreVto <> 1 Then
                        Rem     WFechaElaboracion = ""
                        Rem   End If
                        Rem fin by nan
                        
                        Rem
                        Rem SI ES COLORANTE NO IMPRIME
                        Rem LA FECHA DE VENCIMIENTO
                        Rem
                        XCodigo = Val(Mid$(ZProducto, 4, 5))
                        XTipoPro = ""
                        If Val(Wempresa) = 1 Then
                            If XCodigo >= 0 And XCodigo <= 999 Then
                                WFechaElaboracion = ""
                                XTipoPro = "CO"
                                    Else
                                If XCodigo >= 11000 And XCodigo <= 12999 Then
                                    WFechaElaboracion = ""
                                    XTipoPro = "CO"
                                        Else
                                    XTipoPro = ""
                                End If
                            End If
                        End If
                        
                        
                        
                        
                        spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
                            rstTerminado.Close
                        End If
                            
                        ZCliente = UCase(ZCliente)
                        ZArticulo = UCase(ZArticulo)
                        ZClave = ZCliente + ZArticulo
        
                        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPrecios.RecordCount > 0 Then
                            ZDesArticulo = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
                            rstPrecios.Close
                        End If
                        
                        Call Conecta_Empresa
                                
                        ZSql = "DELETE Certificado"
                        spCertificado = ZSql
                        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        LugarMetodo = 0
                                
                        For CiclaMetodo = 1 To 10
                                
                            If ZOpcion(CiclaMetodo) = 1 Then
                                
                                LugarMetodo = LugarMetodo + 1
                                    
                                ZOrden = ""
                                ZClave1 = ZLote
                                Call Ceros(ZClave1, 6)
                                ZClave2 = Str$(LugarMetodo)
                                Call Ceros(ZClave2, 2)
                                ZClave = ZClave1 + ZClave2
                                ZMetodo = ZEnsayo(CiclaMetodo)
                                
                                Rem If Val(ZStd(CiclaMetodo, 3)) <> 0 And Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                If Val(ZStd(CiclaMetodo, 3)) <> 0 Or Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                    
                                    If Idioma.ListIndex = 0 Then
                                        ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo)) + " " + Left$(ZStd(CiclaMetodo, 1), 50)
                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                            Else
                                        ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo)) + " " + Left$(ZStd(CiclaMetodo, 5), 50)
                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 6), 50)
                                    End If
                                    
                                        Else
                                        
                                    If Idioma.ListIndex = 0 Then
                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                            Else
                                        ZValorNormalI = Left$(ZStd(CiclaMetodo, 5), 50)
                                        ZValorNormalII = Left$(ZStd(CiclaMetodo, 6), 50)
                                    End If
                                    
                                End If
                                
                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 80)
                                If Idioma.ListIndex = 1 Then
                                    If UCase(Trim(ZValorPartidaI)) = "CUMPLE" Then
                                        ZValorPartidaI = "OK"
                                    End If
                                End If
                                
                                If Tipo.ListIndex = 2 Then
                                    ZValorNormalI = Trim(ZValorNormalI)
                                    ZValorNormalII = Trim(ZValorNormalII)
                                    ZValorPartidaI = Trim(ZValorPartidaI)
                                        Else
                                    ZValorNormalI = Trim(ZValorNormalI)
                                    ZCanti = 80 - Len(ZValorNormalI)
                                    ZCanti = Int(ZCanti / 2)
                                    ZValorNormalI = Space$(ZCanti) + ZValorNormalI
                                    
                                    ZValorNormalII = Trim(ZValorNormalII)
                                    ZCanti = 80 - Len(ZValorNormalII)
                                    ZCanti = Int(ZCanti / 2)
                                    ZValorNormalII = Space$(ZCanti) + ZValorNormalII
                                    
                                    ZValorPartidaI = Trim(ZValorPartidaI)
                                    ZCanti = 80 - Len(ZValorPartidaI)
                                    ZCanti = Int(ZCanti / 2)
                                    ZValorPartidaI = Space$(ZCanti) + ZValorPartidaI
                                End If
                                
                                ZValorPartidaII = ""
                                ZObservacionesI = ""
                                ZObservacionesII = ""
                                ZObservacionesIII = "Version " + ZVersion
                                ZObservacionesIV = ""
                                ZObservacionesV = ""
                                ZObservacionesVI = ""
                                If Val(Wempresa) = 1 Then
                                    ZEmpresa = "Surfactan S.A."
                                        Else
                                    ZEmpresa = "Pellital S.A."
                                End If
                                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                ZFechaII = WFechaElaboracion
                                
                                Rem by nan 10-8-2015 para ingles
                                Rem  fechaIng = Left$(Date$, 2) + "/" + Mid$(Date$, 4, 2) + Right$(Date$, 4)
                                Rem    fechaing = Format(Date, "dddd")


                               
                                
                                
                                ZExamen = ZDescri(CiclaMetodo)
                                ZExamenII = ""
                                
                                ZHasta = Len(Trim(ZExamen))
                                If ZHasta > 25 Then
                                    For Cicla = ZHasta To 1 Step -1
                                        If Mid(ZExamen, Cicla, 1) = Space(1) Then
                                            ZExamenII = Mid(ZExamen, Cicla - Desde + 1, 25)
                                            ZExamen = Mid(ZExamen, 1, Cicla - Desde)
                                            Exit For
                                        End If
                                    Next Cicla
                                End If
                                        
                                ZSql = ""
                                ZSql = ZSql + "INSERT INTO Certificado ("
                                ZSql = ZSql + "Clave ,"
                                ZSql = ZSql + "Partida ,"
                                ZSql = ZSql + "Renglon ,"
                                ZSql = ZSql + "Razon ,"
                                ZSql = ZSql + "Orden ,"
                                ZSql = ZSql + "Terminado ,"
                                ZSql = ZSql + "Descripcion ,"
                                ZSql = ZSql + "Fecha ,"
                                ZSql = ZSql + "FechaII ,"
                                ZSql = ZSql + "Cantidad ,"
                                ZSql = ZSql + "Examen ,"
                                ZSql = ZSql + "ExamenII ,"
                                ZSql = ZSql + "ValorPartidaI ,"
                                ZSql = ZSql + "ValorPartidaII ,"
                                ZSql = ZSql + "ValorNormalI ,"
                                ZSql = ZSql + "ValorNormalII ,"
                                ZSql = ZSql + "Observaciones1 ,"
                                ZSql = ZSql + "Observaciones2 ,"
                                ZSql = ZSql + "Observaciones3 ,"
                                ZSql = ZSql + "Observaciones4 ,"
                                ZSql = ZSql + "Observaciones5 ,"
                                ZSql = ZSql + "Observaciones6 ,"
                                ZSql = ZSql + "Metodo ,"
                                Rem  ZSql = ZSql + "fechaing ,"
                                ZSql = ZSql + "Empresa )"
                                ZSql = ZSql + "Values ("
                                ZSql = ZSql + "'" + ZClave + "',"
                                ZSql = ZSql + "'" + ZLote + "',"
                                ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                ZSql = ZSql + "'" + ZRazon + "',"
                                ZSql = ZSql + "'" + ZOrden + "',"
                                ZSql = ZSql + "'" + ZArticulo + "',"
                                ZSql = ZSql + "'" + ZDesArticulo + "',"
                                ZSql = ZSql + "'" + ZFecha + "',"
                                ZSql = ZSql + "'" + ZFechaII + "',"
                                ZSql = ZSql + "'" + ZCantidad + "',"
                                ZSql = ZSql + "'" + ZExamen + "',"
                                ZSql = ZSql + "'" + ZExamenII + "',"
                                ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                ZSql = ZSql + "'" + ZValorNormalI + "',"
                                ZSql = ZSql + "'" + ZValorNormalII + "',"
                                ZSql = ZSql + "'" + ZObservacionesI + "',"
                                ZSql = ZSql + "'" + ZObservacionesII + "',"
                                ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                ZSql = ZSql + "'" + ZObservacionesV + "',"
                                ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                ZSql = ZSql + "'" + ZMetodo + "',"
                                Rem  ZSql = ZSql + "'" + fechaing + "',"
                                ZSql = ZSql + "'" + ZEmpresa + "')"
            
                                spCertificado = ZSql
                                Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                        
                            End If
                                            
                        Next CiclaMetodo
                            
                        Do
                            
                            If LugarMetodo = 10 Then
                                Exit Do
                            End If
                                
                            LugarMetodo = LugarMetodo + 1
                                    
                            ZOrden = ""
                            ZClave1 = ZLote
                            Call Ceros(ZClave1, 6)
                            ZClave2 = Str$(LugarMetodo)
                            Call Ceros(ZClave2, 2)
                            ZClave = ZClave1 + ZClave2
                            ZMetodo = ""
                            ZExamen = ""
                            ZValorNormalI = ""
                            ZValorNormalII = ""
                            ZValorPartidaI = ""
                            ZValorPartidaII = ""
                            ZObservacionesI = ""
                            ZObservacionesII = ""
                            ZObservacionesIII = "Version " + ZVersion
                            ZObservacionesIV = ""
                            ZObservacionesV = ""
                            ZObservacionesVI = ""
                            If Val(Wempresa) = 1 Then
                                ZEmpresa = "Surfactan S.A."
                                    Else
                                ZEmpresa = "Pellital S.A."
                            End If
                            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            ZFechaII = WFechaElaboracion
                            ZExamenII = ""
                                        
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Certificado ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Partida ,"
                            ZSql = ZSql + "Renglon ,"
                            ZSql = ZSql + "Razon ,"
                            ZSql = ZSql + "Orden ,"
                            ZSql = ZSql + "Terminado ,"
                            ZSql = ZSql + "Descripcion ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "FechaII ,"
                            ZSql = ZSql + "Cantidad ,"
                            ZSql = ZSql + "Examen ,"
                            ZSql = ZSql + "ValorPartidaI ,"
                            ZSql = ZSql + "ValorPartidaII ,"
                            ZSql = ZSql + "ValorNormalI ,"
                            ZSql = ZSql + "ValorNormalII ,"
                            ZSql = ZSql + "Observaciones1 ,"
                            ZSql = ZSql + "Observaciones2 ,"
                            ZSql = ZSql + "Observaciones3 ,"
                            ZSql = ZSql + "Observaciones4 ,"
                            ZSql = ZSql + "Observaciones5 ,"
                            ZSql = ZSql + "Observaciones6 ,"
                            ZSql = ZSql + "Metodo ,"
                            Rem   ZSql = ZSql + "fechaing ,"
                            ZSql = ZSql + "Empresa )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZClave + "',"
                            ZSql = ZSql + "'" + ZLote + "',"
                            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                            ZSql = ZSql + "'" + ZRazon + "',"
                            ZSql = ZSql + "'" + ZOrden + "',"
                            ZSql = ZSql + "'" + ZArticulo + "',"
                            ZSql = ZSql + "'" + ZDesArticulo + "',"
                            ZSql = ZSql + "'" + ZFecha + "',"
                            ZSql = ZSql + "'" + ZFechaII + "',"
                            ZSql = ZSql + "'" + ZCantidad + "',"
                            ZSql = ZSql + "'" + ZExamen + "',"
                            ZSql = ZSql + "'" + ZValorPartidaI + "',"
                            ZSql = ZSql + "'" + ZValorPartidaII + "',"
                            ZSql = ZSql + "'" + ZValorNormalI + "',"
                            ZSql = ZSql + "'" + ZValorNormalII + "',"
                            ZSql = ZSql + "'" + ZObservacionesI + "',"
                            ZSql = ZSql + "'" + ZObservacionesII + "',"
                            ZSql = ZSql + "'" + ZObservacionesIII + "',"
                            ZSql = ZSql + "'" + ZObservacionesIV + "',"
                            ZSql = ZSql + "'" + ZObservacionesV + "',"
                            ZSql = ZSql + "'" + ZObservacionesVI + "',"
                            ZSql = ZSql + "'" + ZMetodo + "',"
                            Rem    ZSql = ZSql + "'" + fechaing + "',"
                            ZSql = ZSql + "'" + ZEmpresa + "')"
            
                            spCertificado = ZSql
                            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                
                        Loop
                                
                        Listado.WindowTitle = "Certificado de Analisis"
                        Listado.WindowTop = 0
                        Listado.WindowLeft = 0
                        Listado.WindowWidth = Screen.Width
                        Listado.WindowHeight = Screen.Height
        
                        Listado.Destination = 1
                        Rem Listado.Destination = 0
                                
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            If Idioma.ListIndex = 0 Then
                                Listado.ReportFileName = "CertificadoNuevo.rpt"
                                    Else
                                Listado.ReportFileName = "CertificadoNuevoIngles.rpt"
                            End If
                            If Tipo.ListIndex = 2 Then
                                Listado.ReportFileName = "CertificadoNuevoWord.rpt"
                            End If
                                Else
                            Listado.ReportFileName = "CertificadonuevoPelli.rpt"
                        End If
                                    
                        DbConnect = db.Connect
                        DSQ = getDatabase(DbConnect)
        
                        Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                        + "From " _
                                        + DSQ + ".dbo.Certificado Certificado " _
                                        + "Where " _
                                        + "Certificado.Partida >= 0 AND " _
                                        + "Certificado.Partida <= 999999"
        
                        Listado.Connect = Connect()
                        
                        If Tipo.ListIndex = 0 Then
                            Listado.Destination = 1
                                Else
                            Listado.Destination = 0
                        End If
                        
                        Listado.Action = 1
                                
                    End If
                          
                End If
                
            End If
                
        Next ZCiclo
        
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
    End If
    
    Call Conecta_Empresa

End Sub





